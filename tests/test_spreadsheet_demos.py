import io
import unittest
import zipfile
from pathlib import Path

from lxml import etree
import xlrd

import flask_app


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = PROJECT_ROOT / 'spreadsheet_demos'
VALIDATION_EXAMPLE = {
    'filename': 'mods_default_tiff_images_validation_examples.xlsx',
    'profile': 'modsprofile',
    'sheet': 'tiff_images',
}

DEMO_SPREADSHEETS = [
    {
        'filename': 'mods_default_basic.xlsx',
        'profile': 'modsprofile',
        'sheet': 'basic_records',
        'zip_filename': 'basic_records.zip',
        'generated_files': [
            'demo_archival_item_001.mods.xml',
            'demo_archival_item_002.mods.xml',
        ],
        'alt_text_by_file': {
            'demo_archival_item_002.mods.xml': 'Front view of a campus building with trees and a walkway.',
        },
    },
    {
        'filename': 'mods_default_repeating_fields.xlsx',
        'profile': 'modsprofile',
        'sheet': 'repeating_fields',
        'zip_filename': 'repeating_fields.zip',
        'generated_files': [
            'demo_repeating_001.mods.xml',
            'demo_repeating_002.mods.xml',
        ],
    },
    {
        'filename': 'mods_default_tiff_images_basic.xlsx',
        'profile': 'modsprofile',
        'sheet': 'tiff_images',
        'zip_filename': 'tiff_images.zip',
        'generated_files': [
            'demo_image_0001.mods.xml',
            'demo_image_0002.mods.xml',
        ],
        'alt_text_by_file': {
            'demo_image_0001.mods.xml': (
                'Black and white photograph of a campus building entrance with steps and columns.'
            ),
            'demo_image_0002.mods.xml': (
                'Black and white portrait of a person seated beside a table with papers.'
            ),
        },
    },
    {
        'filename': 'mods_john_nicholas_brown_center_syllabi_basic.xlsx',
        'profile': 'jnbcsyllabi',
        'sheet': 'syllabi',
        'zip_filename': 'syllabi.zip',
        'generated_files': [
            'jnbc_demo_syllabus_001.mods.xml',
            'jnbc_demo_syllabus_002.mods.xml',
        ],
    },
    {
        'filename': 'mods_music_theses_basic.xlsx',
        'profile': 'musictheses',
        'sheet': 'music_theses',
        'zip_filename': 'music_theses.zip',
        'generated_files': [
            'bdrdemo_music_001.mods.xml',
            'bdrdemo_music_002.mods.xml',
        ],
    },
]


class TestSpreadsheetDemos(unittest.TestCase):

    def setUp(self):
        flask_app.app.config['TESTING'] = True
        self.client = flask_app.app.test_client()

    def test_demo_spreadsheets_generate_mods_zip_downloads(self):
        """
        Checks that each demo spreadsheet can be uploaded through the MODS Maker route.
        """
        for demo in DEMO_SPREADSHEETS:
            with self.subTest(filename=demo['filename']):
                workbook_path = DEMO_DIR / demo['filename']
                response = self.client.post(
                    '/modsmaker/%s' % demo['profile'],
                    data={
                        'input_file': (io.BytesIO(workbook_path.read_bytes()), demo['filename']),
                        'sheetlist': demo['sheet'],
                    },
                    content_type='multipart/form-data',
                )

                self.assertEqual(200, response.status_code)
                self.assertEqual(
                    'attachment; filename=%s' % demo['zip_filename'],
                    response.headers['Content-Disposition'],
                )
                self.assert_generated_zip_matches_demo(response.data, demo)
                self.assert_demo_alt_text_values_are_within_limit(workbook_path, demo['sheet'])

    def test_validation_example_spreadsheet_returns_preview_errors(self):
        """
        Checks that the validation example workbook demonstrates expected preview errors.
        """
        workbook_path = DEMO_DIR / VALIDATION_EXAMPLE['filename']
        response = self.client.post(
            '/modsmaker/getpreview',
            data={
                'xlsx_file': (io.BytesIO(workbook_path.read_bytes()), VALIDATION_EXAMPLE['filename']),
                'data': '{"sheetname": "tiff_images", "profile": "modsprofile", "globalconditions": {}}',
            },
            content_type='multipart/form-data',
        )

        errors = response.get_json()['errors']

        self.assertEqual(200, response.status_code)
        self.assertEqual(3, len(errors))
        self.assertEqual([2, 4, 5], [error['spreadsheet_row'] for error in errors])
        self.assertEqual(['required', 'maxchars', 'required'], [error['type'] for error in errors])

    def test_validation_example_spreadsheet_does_not_download_zip(self):
        """
        Checks that the validation example workbook is intentionally invalid for ZIP download.
        """
        workbook_path = DEMO_DIR / VALIDATION_EXAMPLE['filename']
        response = self.client.post(
            '/modsmaker/%s' % VALIDATION_EXAMPLE['profile'],
            data={
                'input_file': (io.BytesIO(workbook_path.read_bytes()), VALIDATION_EXAMPLE['filename']),
                'sheetlist': VALIDATION_EXAMPLE['sheet'],
            },
            content_type='multipart/form-data',
        )

        self.assertEqual(200, response.status_code)
        self.assertNotIn('Content-Disposition', response.headers)
        self.assertIn(b'Row 2, column', response.data)
        self.assertIn(b'Row 4, column', response.data)
        self.assertIn(b'Row 5, column', response.data)
        self.assertIn(b'noteImageAltText', response.data)

    def test_validation_example_spreadsheet_downloads_zip_when_validation_is_not_enforced(self):
        """
        Checks that the validation example workbook can still generate a ZIP when enforcement is disabled.
        """
        workbook_path = DEMO_DIR / VALIDATION_EXAMPLE['filename']
        response = self.client.post(
            '/modsmaker/%s' % VALIDATION_EXAMPLE['profile'],
            data={
                'input_file': (io.BytesIO(workbook_path.read_bytes()), VALIDATION_EXAMPLE['filename']),
                'sheetlist': VALIDATION_EXAMPLE['sheet'],
                'enforce_validations': 'false',
            },
            content_type='multipart/form-data',
        )

        self.assertEqual(200, response.status_code)
        self.assertEqual(
            'attachment; filename=%s.zip' % VALIDATION_EXAMPLE['sheet'],
            response.headers['Content-Disposition'],
        )

    def assert_generated_zip_matches_demo(self, zip_bytes, demo):
        """
        Checks that a generated demo ZIP contains parseable MODS files with expected names.
        """
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zip_file:
            self.assertEqual(demo['generated_files'], zip_file.namelist())
            namespaces = {'mods': 'http://www.loc.gov/mods/v3'}

            for generated_file in demo['generated_files']:
                root = etree.fromstring(zip_file.read(generated_file))
                self.assertEqual('{http://www.loc.gov/mods/v3}mods', root.tag)
                expected_alt_text = demo.get('alt_text_by_file', {}).get(generated_file)
                if expected_alt_text:
                    self.assertEqual(
                        [expected_alt_text],
                        root.xpath('mods:note[@type="image_accessibility_alt_text"]/text()', namespaces=namespaces),
                    )

    def assert_demo_alt_text_values_are_within_limit(self, workbook_path, sheet_name):
        """
        Checks that demo workbook alt text values stay within the enforced limit.
        """
        workbook = xlrd.open_workbook(str(workbook_path))
        sheet = workbook.sheet_by_name(sheet_name)
        headers = [sheet.cell_value(0, column_index) for column_index in range(sheet.ncols)]

        if 'noteImageAltText' not in headers:
            return

        alt_text_column = headers.index('noteImageAltText')
        for row_index in range(1, sheet.nrows):
            alt_text = sheet.cell_value(row_index, alt_text_column).strip()
            self.assertLessEqual(len(alt_text), 250)


if __name__ == '__main__':
    unittest.main()
