import os
import zipfile
from StringIO import StringIO
from copy import deepcopy
from lxml import etree


class ExcelXMLMangler(object):
    NAMESPACES = {'main':
                      "http://schemas.openxmlformats.org/"\
                      "spreadsheetml/2006/main",
                  'pack_rel':
                      'http://schemas.openxmlformats.org/'\
                      'package/2006/relationships',
                  'doc_rel':
                      'http://schemas.openxmlformats.org/'\
                      'officeDocument/2006/relationships'}

    def add_string(self, row, sheet, sheet_filename, hook, data):
        string_id = self.get_shared_string_ref(data)
        row.xpath('//main:c[main:v=\'%s\']/main:v' % hook,
                  namespaces=self.NAMESPACES)[0].text = str(string_id)

    def add_url(self, row, sheet, sheet_filename, hook, data):
        """
        We modify the sheet without updating caches. We expect our caller to
        do it!
        """
        url = data[0]
        string_id = self.get_shared_string_ref(data[1])
        column = row.xpath('//main:c[main:v=\'%s\']/main:v' % hook,
                           namespaces=self.NAMESPACES)[0]
        column.text = str(string_id)
        rel_name = self.calc_rel_file(sheet_filename)
        rid = self.new_relation(rel_name,
                                {'Type': 'http://schemas.openxmlformats.org'
                                 '/officeDocument/2006/relationships/'
                                 'hyperlink', 'Target': url,
                                 'TargetMode': 'External'})
        column_id = column.getparent().attrib['r']
        new_node = etree.fromstring('<hyperlink xmlns="%s" xmlns:r="%s"'
                                    'ref="%s" r:id="%s" />' %
                                    (self.NAMESPACES['main'],
                                     self.NAMESPACES['doc_rel'], column_id,
                                     rid))
        sheet.xpath('//main:hyperlinks', namespaces=self.NAMESPACES)[0]\
            .append(new_node)

    def __init__(self, filename):
        self.fakefile = StringIO(file(filename).read())
        zfile = zipfile.ZipFile(self.fakefile, 'r')
        self.files = {}
        for info in zfile.filelist:
            self.files[info.filename] = zfile.open(info.filename).read()
        self.sheets = {}
        for sheet_info in self.workbook_xml('//main:sheet'):
            name = sheet_info.attrib['name']
            relation_id = sheet_info.attrib['{%s}id' % \
                                                self.NAMESPACES['doc_rel']]
            sheet_id = sheet_info.attrib['sheetId']
            sheet_file = self.files[os.path.join('xl',
                self.xl_relationships[relation_id]['Target'])]
            xml = etree.fromstring(sheet_file)
            self.sheets[name] = {'relation_id': relation_id,
                                 'sheet_filename': os.path.join('xl',
                         self.xl_relationships[relation_id]['Target']),
                                 'sheet_id': sheet_id,
                                 'xml': xml}

    def save(self, filename):
        filenames = self.files.keys()
        filenames.sort()
        target = zipfile.ZipFile(filename, 'w')
        for filename in filenames:
            target.writestr(filename, self.files[filename])
        target.close()

    def copySheet(self, source_sheet, target_sheet):
        sheet = self.sheets[source_sheet]
        new_relation_id, new_filename = self.new_sheet_relation()
        new_id = str(reduce(lambda a, b: max(a, int(b['sheet_id'])),
                            self.sheets.values(), 0) + 1)
        new_sheet_xml = etree.fromstring(
            '<main:sheet xmlns:main="%s" xmlns:doc_rel="%s" name="%s" sheetId'\
                '="%s" doc_rel:id="%s" />' % \
                (self.NAMESPACES['main'], self.NAMESPACES['doc_rel'],
                 target_sheet, new_id, new_relation_id))
        self.workbook_xml('main:sheets')[0].append(new_sheet_xml)
        self.files['xl/workbook.xml'] = etree.tostring(\
            self.workbook_xml()[0])
        new_rel_file = self.calc_rel_file(new_filename)
        old_rel_file = self.calc_rel_file(sheet['sheet_filename'])
        self.files[new_rel_file] = self.files[old_rel_file]
        new_xml = deepcopy(sheet['xml'])
        new_xml_str = etree.tostring(new_xml)
        self.files[new_filename] = new_xml_str
        self.sheets[target_sheet] = {'relation_id': new_relation_id,
                                     'sheet_filename': new_filename,
                                     'sheet_id': new_id,
                                     'xml': new_xml}
        del self._xl_relationships

    def moveSheet(self, sheetName, newPos):
        self.sheets[sheetName]['sheetId'] = newPos
        sheet = self.workbook_xml('//main:sheet[@name=\'%s\']' % sheetName)[0]
        sheets = self.workbook_xml('//main:sheets')[0]
        sheet.attrib['sheetId'] = newPos
        children = sheets.getchildren()
        children.sort(lambda a, b: int(a.attrib['sheetId']).__cmp__(
                int(b.attrib['sheetId'])))
        for child in children:
            sheets.remove(child)
            sheets.append(child)
        self.files['xl/workbook.xml'] = etree.tostring(\
            self.workbook_xml()[0])

    def addRows(self, sheetName, head, data):
        hook = self.get_shared_string_ref(head[0][0])
        sheet = self.sheets[sheetName]['xml']
        sheetdata = sheet.xpath('main:sheetData', namespaces=self.NAMESPACES)[0]
        filename = self.sheets[sheetName]['sheet_filename']
        tmpl_row = sheet.xpath('//main:c[main:v=\'%s\']/..' % hook, namespaces=self.NAMESPACES)[0]
        start_row_number = int(tmpl_row.attrib['r'])
        summary_row = tmpl_row.getnext()
        head = map(lambda x:(self.get_shared_string_ref(x[0]), x[1]), head)
        for index, new_row_data in enumerate(data):
            new_row = deepcopy(tmpl_row)
            new_row_number = start_row_number + index
            self.delete_row(sheet, new_row_number)
            self.update_row_number(str(new_row_number), new_row)
            for (old, new) in zip(head, new_row_data):
                column_type = old[1]
                old_data = old[0]
                if column_type == 'string':
                    self.add_string(new_row, sheet, filename, old_data, new)
                elif column_type == 'url':
                    self.add_url(new_row, sheet, filename, old_data, new)
            sheetdata.insert(new_row_number, new_row)
        new_row_number = start_row_number + index + 1
        self.delete_row(sheet, new_row_number)
        self.update_row_number(str(new_row_number), summary_row)
        sheetdata.insert(new_row_number - 1, summary_row)
        self.files[filename] = etree.tostring(sheet)

    def replaceValue(self, sheetName, old, new):
        sheet = self.sheets[sheetName]['xml']
        filename = self.sheets[sheetName]['sheet_filename']
        old = self.get_shared_string_ref(old)
        new = self.get_shared_string_ref(new)
        sheet.xpath('//main:c[main:v=\'%s\']/main:v' % old, namespaces=self.NAMESPACES)[0].text = str(new)
        self.files[filename] = etree.tostring(sheet)

    def delete_row(self, sheet, row_number):
        """
        Delete the row with the given row number from the sheet
        if it exists.
        If it does not exist, do nothing
        """
        row = sheet.xpath('//main:row[@r=\'%s\']' % row_number,
                          namespaces=self.NAMESPACES)
        if len(row) == 1:
            row[0].getparent().remove(row[0])

    def update_row_number(self, new_number, row):
        old_number = row.attrib['r']
        for child in row.getchildren():
            child.attrib['r'] = child.attrib['r'].replace(old_number, new_number)
        row.attrib['r'] = new_number

    def get_shared_string_ref(self, value):
        if not hasattr(self, '_shared_strings'):
            self._shared_strings = etree.fromstring(self.files['xl/sharedStrings.xml']).xpath('//main:t/text()', namespaces=self.NAMESPACES)
        try:
            return self._shared_strings.index(value)
        except ValueError:
            xml = etree.fromstring(self.files['xl/sharedStrings.xml'])
            new_node = etree.fromstring('<si xmlns="%s"><t>%s</t></si>' %
                                        (self.NAMESPACES['main'], value))
            xml.xpath('/main:sst', namespaces=self.NAMESPACES)[0].append(new_node)
            self.files['xl/sharedStrings.xml'] = etree.tostring(xml)
            retval = len(self._shared_strings)
            del self._shared_strings
            return retval

    def calc_rel_file(self, filename):
        splitted = filename.split(os.path.sep)
        return os.path.join(*(splitted[:-1] \
                                  + ['_rels', splitted[-1] + '.rels']))

    def workbook_xml(self, path='/main:workbook'):
        if not hasattr(self, '_workbook_xml'):
            self._workbook_xml = etree.fromstring(\
                self.files['xl/workbook.xml'])
        return self._workbook_xml.xpath(path, namespaces=self.NAMESPACES)

    def workbook_rels_xml(self, path):
        if not hasattr(self, '_workbook_rels_xml'):
            self._workbook_rels_xml = etree.fromstring(\
                self.files['xl/_rels/workbook.xml.rels'])
        return self._workbook_rels_xml.xpath(path, namespaces=self.NAMESPACES)

    @property
    def xl_relationships(self):
        if not hasattr(self, '_xl_relationships'):
            self._xl_relationships = {}
            for rel in self.workbook_rels_xml('//pack_rel:Relationship'):
                self._xl_relationships[rel.attrib['Id']] = rel.attrib
        return self._xl_relationships

    def new_sheet_relation(self):
        type_def = 'http://schemas.openxmlformats.org/'\
            'officeDocument/2006/relationships/worksheet'
        new_filename = 'worksheets/sheet%i.xml' % \
            (reduce(lambda a, b: max(a, int(b['Target'][16:-4])),
                    filter(lambda x: x['Type'] == type_def,
                           self.xl_relationships.values()), 0) + 1)
        attribs = {
            'Type': type_def,
            'Target': new_filename,
            }
        return self.new_relation('xl/_rels/workbook.xml.rels', attribs),\
            'xl/' + new_filename

    def new_relation(self, filename, attribs):
        xml = etree.fromstring(self.files[filename])
        new_attribs = []
        for key, value in attribs.items():
            new_attribs.append('%s="%s"' % (key, value))
        rid = reduce(lambda a, b:max(a,int(b.attrib['Id'][3:])), xml.xpath('//pack_rel:Relationship', namespaces=self.NAMESPACES), 0) + 1
        new_relation_id = 'rId%i' % rid
        xml.xpath('//pack_rel:Relationships', namespaces=self.NAMESPACES)[0].append(\
            etree.fromstring('<Relationship xmlns="%s" Id="%s" %s />' \
             % (self.NAMESPACES['pack_rel'], new_relation_id,
                ' '.join(new_attribs))))
        self.files[filename] = \
            etree.tostring(xml.xpath(\
                '/pack_rel:Relationships', namespaces=self.NAMESPACES)[0])
        return new_relation_id
