import io
import unittest
import zipfile
from pathlib import Path

from lxml import etree

import flask_app


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = PROJECT_ROOT / 'spreadsheet_demos'

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

    def assert_generated_zip_matches_demo(self, zip_bytes, demo):
        """
        Checks that a generated demo ZIP contains parseable MODS files with expected names.
        """
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zip_file:
            self.assertEqual(demo['generated_files'], zip_file.namelist())

            for generated_file in demo['generated_files']:
                root = etree.fromstring(zip_file.read(generated_file))
                self.assertEqual('{http://www.loc.gov/mods/v3}mods', root.tag)


if __name__ == '__main__':
    unittest.main()
