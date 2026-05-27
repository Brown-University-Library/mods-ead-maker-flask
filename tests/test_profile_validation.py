import unittest

from lib import profileValidation


VALIDATIONS = [
    {
        'type': 'maxchars',
        'col': 'imageAccessibilityAltText',
        'maxchars': 250,
        'severity': 'error',
        'message': 'Image accessibility alt text must be 250 characters or fewer.',
    },
    {
        'type': 'required',
        'col': 'imageAccessibilityAltText',
        'severity': 'error',
        'conditions': [
            {'type': 'equals', 'col': 'typeOfResource', 'text': 'still image'},
        ],
        'message': 'Image accessibility alt text is required when typeOfResource is still image.',
    },
]


class TestProfileValidation(unittest.TestCase):

    def test_validate_rows_allows_alt_text_at_character_limit(self):
        """
        Checks that 250-character alt text passes validation.
        """
        row = {
            'typeOfResource': 'still image',
            'imageAccessibilityAltText': 'a' * 250,
        }

        errors = profileValidation.validateRows([row], VALIDATIONS)

        self.assertEqual([], errors)

    def test_validate_rows_blocks_alt_text_over_character_limit(self):
        """
        Checks that alt text longer than 250 characters fails validation.
        """
        row = {
            'typeOfResource': 'text',
            'imageAccessibilityAltText': 'a' * 251,
        }

        errors = profileValidation.validateRows([row], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual('maxchars', errors[0]['type'])
        self.assertEqual(250, errors[0]['limit'])

    def test_validate_rows_requires_alt_text_for_still_image(self):
        """
        Checks that still image rows require alt text.
        """
        errors = profileValidation.validateRows([{'typeOfResource': 'still image'}], VALIDATIONS)

        self.assertEqual(1, len(errors))
        self.assertEqual('required', errors[0]['type'])

    def test_validate_rows_trims_and_ignores_case_for_still_image_condition(self):
        """
        Checks that condition matching trims whitespace and ignores case.
        """
        errors = profileValidation.validateRows([{'typeOfResource': ' Still Image '}], VALIDATIONS)

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
            {'typeOfResource': 'still image|text'},
            {'typeOfResource': 'still image; text'},
        ]

        errors = profileValidation.validateRows(rows, VALIDATIONS)

        self.assertEqual([], errors)


if __name__ == '__main__':
    unittest.main()
