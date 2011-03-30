import os
import unittest2 as unittest
import zipfile


class BaseTests(unittest.TestCase):
    def get_mangler(self):
        from xlsx_filler import ExcelXMLMangler
        return ExcelXMLMangler('test.xlsx')

    def assertZipEquals(self, a, b):
        from lxml import etree
        def beauty_compare(file_ob):
            data = file_ob.read()
            if data.startswith('<?xml'):
                return etree.tostring(etree.fromstring(data)).strip()
            else:
                return data.strip()
        a = zipfile.ZipFile(a)
        b = zipfile.ZipFile(b)
        filename = lambda x: x.filename
        self.assertEquals(map(filename, a.filelist),
                          map(filename, b.filelist))
        self.assertTrue(len(a.filelist) > 0)
        for filename_a, filename_b in zip(a.filelist, b.filelist):
            data_a = beauty_compare(a.open(filename_a))
            data_b = beauty_compare(b.open(filename_b))
            self.assertEquals(data_a, data_b)

    def test_loading(self):
        """
        See if file loading works
        """
        self.get_mangler()
        self.assertTrue(True)

    def test_sheet_copy(self):
        """
        Try to copy a sheet
        """
        # XXX This will fail if the xlsx sheet does not have any link or
        # relation to anything. Test it!
        filler = self.get_mangler()
        filler.copy_sheet('Fancyname', 'Fancycopy')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_sheet_copy.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_sheet_copy.xlsx'), tmpfile)

    def test_delete_sheet(self):
        """
        Try to delete a sheet
        """
        filler = self.get_mangler()
        filler.copy_sheet('Fancyname', 'Fancycopy')
        filler.delete_sheet('Fancyname')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_delete_sheet.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_delete_sheet.xlsx'), tmpfile)

    def test_multiple_sheet_copy(self):
        """
        Try to copy a sheet multiple times.
        """
        filler = self.get_mangler()
        filler.copy_sheet('Fancyname', 'Fancycopy')
        filler.copy_sheet('Fancyname', 'Fancycopy2')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_multiple_sheet_copy.xlsx', 'w').write(
        #    tmpfile.read())
        self.assertZipEquals(file('test_result_multiple_sheet_copy.xlsx'),
                             tmpfile)

    def test_sheet_shuffling(self):
        """
        Try to change the sheet ordering
        """
        filler = self.get_mangler()
        filler.move_sheet('Fancyname', '99')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_sheet_reshuffled.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_sheet_reshuffled.xlsx'),
                             tmpfile)

    def test_add_rows(self):
        """
        Test the core functionality, adding a number of datarows while keeping
        the row immediatelly below the template row
        """
        # XXX test for xlsx sheets without images inside!
        # XXX test for sparse rows!
        filler = self.get_mangler()
        filler.copy_sheet('Fancyname', 'Fancycopy')
        schema = [('field1', 'url'), ('field2', 'string'),
                  ('field3', 'string')]
        data = [(('http://wwww.example.com', 'link'), 'example', 'ex1', 'ex2'),
                (('http://www.example.com/2', 'link2'), 'example', 'ex3', 'ex4')]
        filler.add_rows('Fancycopy', schema, data)
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_rows_added.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_rows_added.xlsx'), tmpfile)

    def test_add_many_rows(self):
        """
        Test that adding lots of rows will not raise errors
        """
        # XXX test for xlsx sheets without images inside!
        # XXX test for sparse rows!
        filler = self.get_mangler()
        filler.copy_sheet('Fancyname', 'Fancycopy')
        schema = [('field1', 'url'), ('field2', 'string'),
                  ('field3', 'string')]
        data = [(('http://wwww.example.com', 'link'), 'example', 'ex1', 'ex2')
                for ignore in range(100)]

        filler.add_rows('Fancycopy', schema, data)
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_rows_added_many_rows.xlsx',
        #     'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_rows_added_many_rows.xlsx'),
                             tmpfile)

    def test_add_rows_to_copied_sheet(self):
        """
        """
        filler = self.get_mangler()
        schema = [('field1', 'url'), ('field2', 'string'),
                  ('field3', 'string')]
        data = [(('http://wwww.example.com', 'link'), 'example', 'ex1', 'ex2'),
                (('http://www.example.com/2', 'link2'), 'example', 'ex3', 'ex4')]
        filler.add_rows('Fancyname', schema, data)
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_rows_added_to_copy.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_rows_added_to_copy.xlsx'),
                             tmpfile)

    def test_rename_single_column(self):
        """
        Test the renaming of a single column
        """
        filler = self.get_mangler()
        filler.replace_value('Fancyname', '<examplereplacement>', 'no example')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_replacement.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_replacement.xlsx'), tmpfile)
