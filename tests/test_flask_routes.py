import io
import json
import unittest
from unittest.mock import patch

import xlsxwriter

import flask_app


def make_xlsx_file():
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Records')
    worksheet.write(0, 0, 'identifierFileName')
    worksheet.write(0, 1, 'fileTitle')
    worksheet.write(1, 0, 'sample-record')
    worksheet.write(1, 1, 'Sample title')
    workbook.close()
    output.seek(0)
    return output


class TestFlaskRoutes(unittest.TestCase):

    def setUp(self):
        flask_app.app.config['TESTING'] = True
        self.client = flask_app.app.test_client()

    def test_modsmaker_redirects_to_default_profile(self):
        """
        Checks that the default MODS Maker route redirects to the default profile.
        """
        response = self.client.get('/modsmaker')

        self.assertEqual(302, response.status_code)
        self.assertIn('/modsmaker/modsprofile', response.headers['Location'])

    def test_modsmaker_profile_get_returns_success(self):
        """
        Checks that the default MODS Maker profile page renders.
        """
        response = self.client.get('/modsmaker/modsprofile')

        self.assertEqual(200, response.status_code)

    def test_profile_list_returns_success_and_known_profile(self):
        """
        Checks that the profile list page renders known profiles.
        """
        response = self.client.get('/profiles/')

        self.assertEqual(200, response.status_code)
        self.assertIn(b'modsprofile', response.data)

    def test_forms_list_returns_success_and_known_profile(self):
        """
        Checks that the forms list page renders known profiles.
        """
        response = self.client.get('/forms/')

        self.assertEqual(200, response.status_code)
        self.assertIn(b'modsprofile', response.data)

    def test_resources_returns_success(self):
        """
        Checks that the resources page renders.
        """
        response = self.client.get('/resources')

        self.assertEqual(200, response.status_code)

    def test_unknown_route_redirects_to_default_modsmaker_profile(self):
        """
        Checks that unknown routes use the default 404 redirect.
        """
        response = self.client.get('/missing-route')

        self.assertEqual(302, response.status_code)
        self.assertIn('/modsmaker/modsprofile', response.headers['Location'])

    def test_process_file_upload_returns_filename_and_sheet_names(self):
        """
        Checks that uploaded workbooks return JSON sheet metadata.
        """
        response = self.client.post(
            '/processfileupload',
            data={'xlsx_file': (make_xlsx_file(), 'records.xlsx')},
            content_type='multipart/form-data',
        )

        self.assertEqual(200, response.status_code)
        self.assertEqual({'filename': 'records.xlsx', 'sheetnames': ['Records']}, response.get_json())

    def test_modsmaker_get_preview_returns_json_preview_text(self):
        """
        Checks that MODS preview route returns the generated preview text as JSON.
        """
        with patch('flask_app.fileSupport.getPreview', return_value='preview text') as mock_get_preview:
            response = self.client.post(
                '/modsmaker/getpreview',
                data={
                    'xlsx_file': (make_xlsx_file(), 'records.xlsx'),
                    'data': json.dumps({
                        'sheetname': 'Records',
                        'profile': 'modsprofile',
                        'globalconditions': {'includeBrownDefaults': True},
                    }),
                },
                content_type='multipart/form-data',
            )

        self.assertEqual(200, response.status_code)
        self.assertEqual('preview text', response.get_json())
        mock_get_preview.assert_called_once()

    def test_modsmaker_post_with_non_xlsx_returns_error(self):
        """
        Checks that non-XLSX MODS uploads render an error response.
        """
        response = self.client.post(
            '/modsmaker/modsprofile',
            data={'input_file': (io.BytesIO(b'not a workbook'), 'records.txt')},
            content_type='multipart/form-data',
        )

        self.assertEqual(200, response.status_code)
        self.assertIn(b'Please go back and select a .XLSX Excel file to proceed.', response.data)

    def test_form_preview_returns_json_preview_text(self):
        """
        Checks that profile form previews return generated preview text as JSON.
        """
        with patch('flask_app.fileSupport.createPreviewFromRows', return_value='preview text') as mock_preview:
            response = self.client.post(
                '/forms/profile/modsprofile/preview',
                data={'identifierFileName': 'sample-record', 'fileTitle': 'Sample title'},
            )

        self.assertEqual(200, response.status_code)
        self.assertEqual('preview text', response.get_json())
        mock_preview.assert_called_once()


if __name__ == '__main__':
    unittest.main()
