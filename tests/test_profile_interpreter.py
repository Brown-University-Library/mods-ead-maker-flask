import unittest

from lxml import etree

from lib import fileSupport
from lib import profileInterpreter


class TestProfileInterpreterParsing(unittest.TestCase):

    def test_normalize_string_removes_markup_pipes_and_repeated_whitespace(self):
        """
        Checks that normalized strings remove selected markup and collapse whitespace.
        """
        normalized = profileInterpreter.normalizeString(' <title>My\n  Title</title>| ')

        self.assertEqual('My Title', normalized)

    def test_get_value_uri_extracts_http_uri(self):
        """
        Checks that an HTTP URI is extracted from entry text.
        """
        uri = profileInterpreter.getValueUri('Example name https://example.org/authority/123')

        self.assertEqual('https://example.org/authority/123', uri)

    def test_get_additional_values_parses_bracketed_yaml(self):
        """
        Checks that bracketed YAML additions become entry metadata.
        """
        values, raw_values = profileInterpreter.getAdditionalValues('Name [displayForm: Jane Doe]')

        self.assertEqual({'entry.displayForm': 'Jane Doe'}, values)
        self.assertEqual(['displayForm: Jane Doe'], raw_values)

    def test_get_name_date_role_from_entry_handles_creator_names(self):
        """
        Checks that creator names receive the Creator role.
        """
        name, date, role = profileInterpreter.getNameDateRoleFromEntry('Doe, Jane, 1970-', 'nameCreator')

        self.assertEqual('Doe, Jane', name)
        self.assertEqual('1970-', date)
        self.assertEqual('Creator', role)

    def test_get_name_date_role_from_entry_handles_other_names(self):
        """
        Checks that other-name entries use the final comma part as role.
        """
        name, date, role = profileInterpreter.getNameDateRoleFromEntry(
            'Doe, Jane, 1970-, Photographer',
            'nameOther',
        )

        self.assertEqual('Doe, Jane', name)
        self.assertEqual('1970-', date)
        self.assertEqual('Photographer', role)

    def test_get_metadata_from_entry_combines_name_uri_role_and_additional_values(self):
        """
        Checks that repeating-entry metadata is parsed into expected fields.
        """
        metadata = profileInterpreter.getMetadataFromEntry(
            'Doe, Jane, 1970-, Photographer https://example.org/id [displayForm: Jane Doe]',
            'nameOther',
        )

        self.assertEqual('https://example.org/id', metadata['entry.valueURI'])
        self.assertEqual('Doe, Jane', metadata['entry.name'])
        self.assertEqual('1970-', metadata['entry.date'])
        self.assertEqual('Photographer', metadata['entry.role'])
        self.assertEqual('Jane Doe', metadata['entry.displayForm'])

    def test_get_key_value_from_entry_splits_on_colon(self):
        """
        Checks that key-value entries split into stripped key and value text.
        """
        key, value = profileInterpreter.getKeyValueFromEntry('local: MS-1')

        self.assertEqual('local', key)
        self.assertEqual('MS-1', value)


class TestProfileInterpreterXml(unittest.TestCase):

    def test_profile_convert_row_to_xml_string_creates_mods_root_and_title(self):
        """
        Checks that the default MODS profile creates parseable MODS XML with a title.
        """
        profile = profileInterpreter.Profile(
            'profiles/modsprofile.yaml',
            globalConditions={'includeBrownDefaults': True, 'includePreferredCitation': True},
        )

        xml_string = profile.convertRowToXmlString({
            'identifierFileName': 'sample-record',
            'fileTitle': 'Sample title',
        })

        root = etree.fromstring(xml_string.encode('utf-8'))
        namespaces = {'mods': 'http://www.loc.gov/mods/v3'}

        self.assertEqual('{http://www.loc.gov/mods/v3}mods', root.tag)
        self.assertEqual(['Sample title'], root.xpath('mods:titleInfo/mods:title/text()', namespaces=namespaces))

    def test_profile_convert_row_to_xml_string_skips_rows_with_skip_column(self):
        """
        Checks that rows with configured skip columns do not generate XML.
        """
        profile = profileInterpreter.Profile('profiles/modsprofile.yaml')

        xml_string = profile.convertRowToXmlString({
            'identifierFileName': 'sample-record',
            'fileTitle': 'Sample title',
            'Ignore': 'yes',
        })

        self.assertIsNone(xml_string)

    def test_profile_convert_row_to_xml_string_creates_image_accessibility_note(self):
        """
        Checks that image accessibility alt text is written as a typed MODS note.
        """
        profile = profileInterpreter.Profile('profiles/modsprofile.yaml')

        xml_string = profile.convertRowToXmlString({
            'identifierFileName': 'sample-record',
            'fileTitle': 'Sample title',
            'noteImageAltText': 'Photograph of a campus building entrance.',
        })

        root = etree.fromstring(xml_string.encode('utf-8'))
        namespaces = {'mods': 'http://www.loc.gov/mods/v3'}

        self.assertEqual(
            ['Photograph of a campus building entrance.'],
            root.xpath('mods:note[@type="image_accessibility_alt_text"]/text()', namespaces=namespaces),
        )

    def test_profile_convert_row_to_xml_string_omits_blank_image_accessibility_note(self):
        """
        Checks that blank image accessibility alt text does not leave an empty note.
        """
        profile = profileInterpreter.Profile('profiles/modsprofile.yaml')

        xml_string = profile.convertRowToXmlString({
            'identifierFileName': 'sample-record',
            'fileTitle': 'Sample title',
            'noteImageAltText': '',
        })

        root = etree.fromstring(xml_string.encode('utf-8'))
        namespaces = {'mods': 'http://www.loc.gov/mods/v3'}

        self.assertEqual([], root.xpath('mods:note[@type="image_accessibility_alt_text"]', namespaces=namespaces))

    def test_create_file_from_row_uses_profile_filename_extension_and_xml(self):
        """
        Checks that a row generates a MODS filename and parseable XML.
        """
        xml_string, file_buffer_value, filename = fileSupport.createFileFromRow(
            {'identifierFileName': 'sample-record', 'fileTitle': 'Sample title'},
            0,
            'profiles/modsprofile.yaml',
            {'includeBrownDefaults': True, 'includePreferredCitation': True},
        )

        self.assertEqual('sample-record.mods.xml', filename)
        self.assertEqual(xml_string, file_buffer_value)
        root = etree.fromstring(file_buffer_value.encode('utf-8'))
        self.assertEqual('{http://www.loc.gov/mods/v3}mods', root.tag)

    def test_create_preview_from_rows_includes_filename_and_xml(self):
        """
        Checks that preview output includes generated filenames and XML for rows.
        """
        preview = fileSupport.createPreviewFromRows(
            [{'identifierFileName': 'sample-record', 'fileTitle': 'Sample title'}],
            'profiles/modsprofile.yaml',
            {'includeBrownDefaults': True, 'includePreferredCitation': True},
        )

        self.assertIn('sample-record.mods.xml', preview)
        self.assertIn('<mods:title>Sample title</mods:title>', preview)


if __name__ == '__main__':
    unittest.main()
