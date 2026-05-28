import unittest
from unittest.mock import patch

from lib import profileValidation


ALT_TEXT_REQUIRED_TYPE_OF_RESOURCE_VALUES = [
    'still image',
    'books',
    'images',
    'journals',
    'manuscripts',
    'maps',
    'newspapers',
    'realia',
    'scores',
    'text_resources',
]

VALIDATIONS = [
    {
        'type': 'maxchars',
        'col': 'noteImageAltText',
        'maxchars': 250,
        'severity': 'error',
        'message': 'Image accessibility alt text must be {maxchars} characters or fewer.',
    },
    {
        'type': 'required',
        'col': 'noteImageAltText',
        'severity': 'error',
        'conditions': [
            {
                'type': 'in',
                'col': 'typeOfResource',
                'values': ALT_TEXT_REQUIRED_TYPE_OF_RESOURCE_VALUES,
            },
        ],
        'message': 'Image accessibility alt text is required when typeOfResource requires it.',
    },
]


class TestProfileValidation(unittest.TestCase):

    def setUp(self):
        self.environPatcher = patch.dict('os.environ', {}, clear=True)
        self.dotenvPatcher = patch('lib.profileValidation.dotenv_values', return_value={})
        self.environPatcher.start()
        self.dotenvPatcher.start()

    def tearDown(self):
        self.dotenvPatcher.stop()
        self.environPatcher.stop()

    def test_validate_rows_allows_alt_text_at_character_limit(self):
        """
        Checks that 250-character alt text passes validation.
        """
        row = {
            'typeOfResource': 'still image',
            'noteImageAltText': 'a' * 250,
        }

        errors = profileValidation.validateRows([row], VALIDATIONS)

        self.assertEqual([], errors)

    def test_validate_rows_blocks_alt_text_over_character_limit(self):
        """
        Checks that alt text longer than 250 characters fails validation.
        """
        row = {
            'typeOfResource': 'text',
            'noteImageAltText': 'a' * 251,
        }

        errors = profileValidation.validateRows([row], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual('maxchars', errors[0]['type'])
        self.assertEqual(250, errors[0]['limit'])
        self.assertEqual('Image accessibility alt text must be 250 characters or fewer.', errors[0]['message'])

    def test_validate_rows_uses_dotenv_alt_text_character_limit_override(self):
        """
        Checks that the parent .env alt text limit overrides the profile limit.
        """
        with patch('lib.profileValidation.dotenv_values', return_value={
            'IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS': '100',
        }):
            errors = profileValidation.validateRows([{
                'typeOfResource': 'text',
                'noteImageAltText': 'a' * 101,
            }], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual(100, errors[0]['limit'])
        self.assertEqual('Image accessibility alt text must be 100 characters or fewer.', errors[0]['message'])

    def test_validate_rows_uses_profile_limit_when_dotenv_alt_text_limit_is_invalid(self):
        """
        Checks that invalid .env alt text limits fall back to the profile limit.
        """
        with patch('lib.profileValidation.dotenv_values', return_value={
            'IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS': 'not-a-number',
        }):
            errors = profileValidation.validateRows([{
                'typeOfResource': 'text',
                'noteImageAltText': 'a' * 251,
            }], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual(250, errors[0]['limit'])

    def test_validate_rows_requires_alt_text_for_configured_type_of_resource_values(self):
        """
        Checks that configured typeOfResource values require alt text.
        """
        rows = [{'typeOfResource': value} for value in ALT_TEXT_REQUIRED_TYPE_OF_RESOURCE_VALUES]

        errors = profileValidation.validateRows(rows, VALIDATIONS)

        errorTypes = [error['type'] for error in errors]

        self.assertEqual(len(ALT_TEXT_REQUIRED_TYPE_OF_RESOURCE_VALUES), len(errors))
        self.assertEqual(['required'] * len(ALT_TEXT_REQUIRED_TYPE_OF_RESOURCE_VALUES), errorTypes)

    def test_validate_rows_trims_and_ignores_case_for_type_of_resource_condition(self):
        """
        Checks that condition matching trims whitespace and ignores case.
        """
        errors = profileValidation.validateRows([{'typeOfResource': ' Images '}], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual('required', errors[0]['type'])

    def test_validate_rows_waives_required_rule_for_non_exact_type_of_resource(self):
        """
        Checks that missing, blank, non-image, and multiple-value typeOfResource values do not require alt text.
        """
        rows = [
            {},
            {'typeOfResource': ''},
            {'typeOfResource': 'moving image'},
            {'typeOfResource': 'book'},
            {'typeOfResource': 'still image|text'},
            {'typeOfResource': 'books|text'},
            {'typeOfResource': 'still image; text'},
        ]

        errors = profileValidation.validateRows(rows, VALIDATIONS)

        self.assertEqual([], errors)


if __name__ == '__main__':
    unittest.main()
