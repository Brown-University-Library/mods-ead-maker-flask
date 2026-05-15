import io
import unittest
import zipfile

from lxml import etree
import xlsxwriter

import fileSupport


def make_xlsx_bytes(sheets):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    for sheet_name, rows in sheets:
        worksheet = workbook.add_worksheet(sheet_name)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                worksheet.write(row_index, column_index, value)

    workbook.close()
    return output.getvalue()


class TestFileSupport(unittest.TestCase):

    def test_clean_string_for_filename_removes_invalid_characters(self):
        """
        Checks that characters invalid in filenames are removed.
        """
        filename = fileSupport.cleanStringForFilename('bad<>:"/\\|?*name')

        self.assertEqual('badname', filename)

    def test_get_filename_from_row_uses_cleaned_column_value(self):
        """
        Checks that row filename values are cleaned before use.
        """
        filename = fileSupport.getFilenameFromRow({'identifierFileName': 'box:1/folder?2'}, 7, 'identifierFileName')

        self.assertEqual('box1folder2', filename)

    def test_get_filename_from_row_falls_back_to_default_index(self):
        """
        Checks that missing filename values use the default index name.
        """
        filename = fileSupport.getFilenameFromRow({}, 3, 'identifierFileName')

        self.assertEqual('default3', filename)

    def test_get_sheet_names_from_xlsx_returns_workbook_sheet_names(self):
        """
        Checks that uploaded workbook bytes yield their sheet names.
        """
        workbook_bytes = make_xlsx_bytes([
            ('First', [['title'], ['One']]),
            ('Second', [['title'], ['Two']]),
        ])

        sheet_names = fileSupport.getSheetNamesFromXlsx(workbook_bytes)

        self.assertEqual(['First', 'Second'], sheet_names)

    def test_convert_xlsx_to_dict_list_uses_headers_and_serializes_repeated_headers(self):
        """
        Checks that spreadsheet rows become dictionaries and repeated headers are joined.
        """
        workbook_bytes = make_xlsx_bytes([
            ('Records', [
                ['title', 'subject', 'subject', 'number'],
                ['Item title', 'alpha', 'beta', 4],
            ]),
        ])

        rows = fileSupport.convertXlsxToDictList(workbook_bytes, 'Records')

        self.assertEqual(1, len(rows))
        self.assertEqual('Item title', rows[0]['title'])
        self.assertEqual('alpha|beta', rows[0]['subject'])
        self.assertEqual('4.0', rows[0]['number'])

    def test_create_zip_from_excel_contains_generated_mods_file(self):
        """
        Checks that spreadsheet rows can be generated into a ZIP of MODS XML files.
        """
        workbook_bytes = make_xlsx_bytes([
            ('Records', [
                ['identifierFileName', 'fileTitle'],
                ['sample-record', 'Sample title'],
            ]),
        ])

        zip_bytes, filename = fileSupport.createZipFromExcel(
            workbook_bytes,
            'Records',
            'profiles/modsprofile.yaml',
            {'includeBrownDefaults': True, 'includePreferredCitation': True},
        )

        self.assertEqual('Records.zip', filename)
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zip_file:
            self.assertEqual(['sample-record.mods.xml'], zip_file.namelist())
            xml_bytes = zip_file.read('sample-record.mods.xml')

        root = etree.fromstring(xml_bytes)
        self.assertEqual('{http://www.loc.gov/mods/v3}mods', root.tag)


if __name__ == '__main__':
    unittest.main()
