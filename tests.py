import os
import unittest2 as unittest
import zipfile


class BaseTests(unittest.TestCase):
    def get_mangler(self):
        from xlsx_filler import ExcelXMLMangler
        return ExcelXMLMangler('test.xlsx')

    def assertZipEquals(self, a, b):
        a = zipfile.ZipFile(a)
        b = zipfile.ZipFile(b)
        filename = lambda x: x.filename
        self.assertEquals(map(filename, a.filelist),
                          map(filename, b.filelist))
        self.assertTrue(len(a.filelist) > 0)
        for filename_a, filename_b in zip(a.filelist, b.filelist):
            self.assertEquals(a.open(filename_a).read(),
                              b.open(filename_b).read())

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
        filler.copySheet('Fancyname', 'Fancycopy')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_sheet_copy.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_sheet_copy.xlsx'), tmpfile)

    def test_sheet_shuffling(self):
        """
        Try to change the sheet ordering
        """
        filler = self.get_mangler()
        filler.moveSheet('Fancyname', '99')
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
        filler.copySheet('Fancyname', 'Fancycopy')
        schema = [('field1', 'url'), ('field2', 'string'),
                  ('field3', 'string')]
        data = [(('http://wwww.example.com'), 'example', 'ex1', 'ex2'),
                (('http://www.example.com/2'), 'example', 'ex3', 'ex4')]
        filler.addRows('Fancycopy', schema, data)
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_rows_added.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_rows_added.xlsx'), tmpfile)

    def test_add_rows_to_copied_sheet(self):
        """
        """
        filler = self.get_mangler()
        schema = [('field1', 'url'), ('field2', 'string'),
                  ('field3', 'string')]
        data = [(('http://wwww.example.com'), 'example', 'ex1', 'ex2'),
                (('http://www.example.com/2'), 'example', 'ex3', 'ex4')]
        filler.addRows('Fancyname', schema, data)
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
        filler.replaceValue('Fancyname', '<examplereplacement>', 'no example')
        tmpfile = os.tmpfile()
        filler.save(tmpfile)
        #tmpfile.seek(0)
        #file('test_result_replacement.xlsx', 'w').write(tmpfile.read())
        self.assertZipEquals(file('test_result_replacement.xlsx'), tmpfile)
