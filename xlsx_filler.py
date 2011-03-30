import re
import os
import zipfile
from StringIO import StringIO
from copy import deepcopy
from lxml import etree


NAMESPACES = {'main':
                  "http://schemas.openxmlformats.org/"\
                  "spreadsheetml/2006/main",
              'pack_rel':
                  'http://schemas.openxmlformats.org/'\
                  'package/2006/relationships',
              'doc_rel':
                  'http://schemas.openxmlformats.org/'\
                  'officeDocument/2006/relationships'}


convert_shared_strings_re = re.compile('<v>([0-9]+)</v>')


def convert_shared_strings(xml_stream, shared_string_converter):
    converter = lambda match: '<v>%s</v>' % \
        shared_string_converter(match.group()[3:-4])
    return convert_shared_strings_re.sub(converter, xml_stream)


class ReferenceDataNotFound(Exception):
    def __init__(self, hook, row, shared_string_converter_):
        shared_string_converter = lambda x: shared_string_converter_(int(x))
        msg_tmpl = "Unable to find the referenced string in the "\
            "reference row.\n"\
            "The string to replace seems to exist, but not in "\
            "the reference row.\n"\
            "refence: %s\n"\
            "reference row: %s\n"
        super(ReferenceDataNotFound, self).__init__(msg_tmpl % (
                shared_string_converter(hook), convert_shared_strings(
                    etree.tostring(row), shared_string_converter)))


class XMLThing(object):

    def __init__(self, thing):
        try:
            self.xml = etree.fromstring(thing)
        except:
            self.xml = thing

    def __str__(self):
        return etree.tostring(self.xml)

    def xpath(self, term):
        return self.xml.xpath(term, namespaces=NAMESPACES)

    def xpath1(self, term):
        return self.xpath(term)[0]


class ExcelXMLMangler(object):
    def __init__(self, filename):
        self.fakefile = StringIO(file(filename).read())
        zfile = zipfile.ZipFile(self.fakefile, 'r')
        self.files = {}
        for info in zfile.filelist:
            data = zfile.open(info.filename).read()
            if data.startswith('<?xml'):
                self.files[info.filename] = XMLThing(zfile.open(info.filename)
                                                     .read())
            else:
                self.files[info.filename] = data
        self.sheets = {}
        for sheet_info in self.files['xl/workbook.xml'].xpath('//main:sheet'):
            name = sheet_info.attrib['name']
            relation_id = sheet_info.attrib['{%s}id' % \
                                                NAMESPACES['doc_rel']]
            sheet_id = sheet_info.attrib['sheetId']
            self.sheets[name] = {'relation_id': relation_id,
                                 'sheet_filename': os.path.join('xl',
                         self._xl_relationships[relation_id]['Target']),
                                 'sheet_id': sheet_id}

    def save(self, filename):
        filenames = self.files.keys()
        filenames.sort()
        target = zipfile.ZipFile(filename, 'w')
        for filename in filenames:
            target.writestr(filename, str(self.files[filename]))
        target.close()

    def _add_string(self, row, sheet, sheet_filename, hook, data):
        row = XMLThing(row)
        string_id = self._get_shared_string_ref(data)
        try:
            row.xpath1('//main:c[main:v=\'%s\']/main:v' % hook).text = \
                str(string_id)
        except IndexError:
            raise ReferenceDataNotFound(hook, row, self._get_shared_string)

    def _add_url(self, row, sheet, sheet_filename, hook, data):
        row = XMLThing(row)
        url = data[0]
        string_id = self._get_shared_string_ref(data[1])
        column = row.xpath1('//main:c[main:v=\'%s\']/main:v' % hook)
        column.text = str(string_id)
        rel_name = self._calc_rel_file(sheet_filename)
        rid = self._new_relation(rel_name,
                                {'Type': 'http://schemas.openxmlformats.org'
                                 '/officeDocument/2006/relationships/'
                                 'hyperlink', 'Target': url,
                                 'TargetMode': 'External'})
        column_id = column.getparent().attrib['r']
        new_node = etree.fromstring('<hyperlink xmlns="%s" xmlns:r="%s"'
                                    'ref="%s" r:id="%s" />' %
                                    (NAMESPACES['main'],
                                     NAMESPACES['doc_rel'], column_id,
                                     rid))
        sheet.xpath1('//main:hyperlinks').append(new_node)

    def copy_sheet(self, source_sheet, target_sheet):
        sheet = self.sheets[source_sheet]
        new_relation_id, new_filename = self._new_sheet_relation()
        new_id = str(reduce(lambda a, b: max(a, int(b['sheet_id'])),
                            self.sheets.values(), 0) + 1)
        new_sheet_xml = etree.fromstring(
            '<main:sheet xmlns:main="%s" xmlns:doc_rel="%s" name="%s" sheetId'\
                '="%s" doc_rel:id="%s" />' % \
                (NAMESPACES['main'], NAMESPACES['doc_rel'],
                 target_sheet, new_id, new_relation_id))
        self.files['xl/workbook.xml'].xpath1('main:sheets')\
            .append(new_sheet_xml)
        new_rel_file = self._calc_rel_file(new_filename)
        old_rel_file = self._calc_rel_file(sheet['sheet_filename'])
        self.files[new_rel_file] = self.files[old_rel_file]
        new_xml = deepcopy(self.files[sheet['sheet_filename']])
        self.files[new_filename] = new_xml
        self.sheets[target_sheet] = {'relation_id': new_relation_id,
                                     'sheet_filename': new_filename,
                                     'sheet_id': new_id}

    def delete_sheet(self, sheet_name):
        sheet_data = self.sheets.pop(sheet_name)
        self.files.pop(sheet_data['sheet_filename'])
        sheet = self.files['xl/workbook.xml']\
            .xpath1('//main:sheet[@name="%s"]' % sheet_name)
        sheet.getparent().remove(sheet)

    def move_sheet(self, sheetName, newPos):
        self.sheets[sheetName]['sheetId'] = newPos
        wb = self.files['xl/workbook.xml']
        sheet = wb.xpath1('//main:sheet[@name=\'%s\']' % sheetName)
        sheets = wb.xpath1('//main:sheets')
        sheet.attrib['sheetId'] = newPos
        children = sheets.getchildren()
        children.sort(lambda a, b: int(a.attrib['sheetId']).__cmp__(
                int(b.attrib['sheetId'])))
        for child in children:
            sheets.remove(child)
            sheets.append(child)

    def add_rows(self, sheetName, head, data):
        hook = self._get_shared_string_ref(head[0][0])
        filename = self.sheets[sheetName]['sheet_filename']
        sheet = self.files[filename]
        sheetdata = sheet.xpath1('main:sheetData')
        tmpl_row = sheet.xpath1('//main:c[main:v=\'%s\']/..' % hook)
        start_row_number = int(tmpl_row.attrib['r'])
        summary_row = tmpl_row.getnext()
        head = map(lambda x: (self._get_shared_string_ref(x[0]), x[1]), head)

        for index, new_row_data in enumerate(data):
            new_row = deepcopy(tmpl_row)
            new_row_number = start_row_number + index
            self.delete_row(sheet, new_row_number)
            self._update_row_number(str(new_row_number), new_row)
            for (old, new) in zip(head, new_row_data):
                column_type = old[1]
                old_data = old[0]
                if column_type == 'string':
                    self._add_string(new_row, sheet, filename, old_data, new)
                elif column_type == 'url':
                    self._add_url(new_row, sheet, filename, old_data, new)
            sheetdata.insert(new_row_number, new_row)
        new_row_number = start_row_number + index + 1
        self.delete_row(sheet, new_row_number)
        self._update_row_number(str(new_row_number), summary_row)
        sheetdata.insert(new_row_number - 1, summary_row)

    def replace_value(self, sheetName, old, new):
        filename = self.sheets[sheetName]['sheet_filename']
        sheet = self.files[filename]
        old = self._get_shared_string_ref(old)
        new = self._get_shared_string_ref(new)
        sheet.xpath1('//main:c[main:v=\'%s\']/main:v' % old).text = str(new)

    def delete_row(self, sheet, row_number):
        """
        Delete the row with the given row number from the sheet
        if it exists.
        If it does not exist, do nothing
        """
        sheet = sheet
        row = sheet.xpath('//main:row[@r=\'%s\']' % row_number)
        if len(row) == 1:
            row[0].getparent().remove(row[0])

    def _update_row_number(self, new_number, row):
        old_number = row.attrib['r']
        for child in row.getchildren():
            child.attrib['r'] = child.attrib['r']\
                .replace(old_number, new_number)
        row.attrib['r'] = new_number

    def _get_shared_string(self, ref_id):
        return self.files['xl/sharedStrings.xml'].xpath('//main:t/text()')

    def _get_shared_string_ref(self, value):
        shared_strings = self.files['xl/sharedStrings.xml']
        shared_strings_list = shared_strings.xpath('//main:t/text()')
        try:
            return shared_strings_list.index(value)
        except ValueError:
            new_node = etree.fromstring('<si xmlns="%s"><t>%s</t></si>' %
                                        (NAMESPACES['main'], value))
            shared_strings.xpath1('/main:sst').append(new_node)
            return len(shared_strings_list)

    def _calc_rel_file(self, filename):
        splitted = filename.split(os.path.sep)
        return os.path.join(*(splitted[:-1] \
                                  + ['_rels', splitted[-1] + '.rels']))

    def _workbook_refs_xml(self, path):
        workbook_rels_xml = etree.fromstring( \
            self.files['xl/_rels/workbook.xml.rels'])
        return workbook_rels_xml.xpath(path, namespaces=NAMESPACES)

    @property
    def _xl_relationships(self):
        xl_relationships = {}
        refs = self.files['xl/_rels/workbook.xml.rels']
        for rel in refs.xpath('//pack_rel:Relationship'):
            xl_relationships[rel.attrib['Id']] = rel.attrib
        return xl_relationships

    def _new_sheet_relation(self):
        type_def = 'http://schemas.openxmlformats.org/'\
            'officeDocument/2006/relationships/worksheet'
        new_filename = 'worksheets/sheet%i.xml' % \
            (reduce(lambda a, b: max(a, int(b['Target'][16:-4])),
                    filter(lambda x: x['Type'] == type_def,
                           self._xl_relationships.values()), 0) + 1)
        attribs = {
            'Type': type_def,
            'Target': new_filename,
            }
        return self._new_relation('xl/_rels/workbook.xml.rels', attribs),\
            'xl/' + new_filename

    def _new_relation(self, filename, attribs):
        xml = self.files[filename]
        new_attribs = []
        for key, value in attribs.items():
            new_attribs.append('%s="%s"' % (key, value))
        rid = reduce(lambda a, b:max(a,int(b.attrib['Id'][3:])),\
                         xml.xpath('//pack_rel:Relationship'), 0) + 1
        new_relation_id = 'rId%i' % rid
        xml.xpath1('//pack_rel:Relationships').append(\
            etree.fromstring('<Relationship xmlns="%s" Id="%s" %s />' \
             % (NAMESPACES['pack_rel'], new_relation_id,
                ' '.join(new_attribs))))
        return new_relation_id
