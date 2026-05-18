# Spreadsheet Demos

These demo spreadsheets contain fictional sample metadata for trying the MODS Maker locally.

To try a spreadsheet upload:
- Start the webapp
- Go to the "route" url listed below -- it's the url-route that selects the mapping-profile
- Upload the spreadsheet, and select the listed sheet
- Preview the output, and optionally download the ZIP of generated `.mods.xml` files

| File | Route | Sheet | Demonstrates |
| --- | --- | --- | --- |
| `mods_default_basic.xlsx` | `/modsmaker/modsprofile` | `basic_records` | Basic MODS fields and download flow |
| `mods_default_repeating_fields.xlsx` | `/modsmaker/modsprofile` | `repeating_fields` | Names, subjects, repeated values, URIs |
| `mods_default_tiff_images_basic.xlsx` | `/modsmaker/modsprofile` | `tiff_images` | Still-image/TIFF-oriented MODS fields |
| `mods_john_nicholas_brown_center_syllabi_basic.xlsx` | `/modsmaker/jnbcsyllabi` | `syllabi` | John Nicholas Brown Center syllabi profile |
| `mods_music_theses_basic.xlsx` | `/modsmaker/musictheses` | `music_theses` | Music thesis metadata profile |

Notes:

- Each workbook has two sample records.
- Generated output is downloaded as a ZIP containing one `.mods.xml` file per generated record.
- Still-image rows include `imageAccessibilityAltText`; the MODS Maker requires that column when `typeOfResource` is `still image`.
- By default, validation errors stop Preview and Download. Uncheck `Enforce validations` to continue processing while showing validation warnings.
- The TIFF image demo uses `identifierFileName` values like `demo_image_0001`, intended to pair conceptually with source files such as `demo_image_0001.tif`.
- The sample metadata is fictional and intended only for demonstration.
- EAD demo spreadsheets are not included yet; they can be added later after a minimal EAD example is confirmed.

## Validation Examples

Use `mods_default_tiff_images_validation_examples.xlsx` with route `/modsmaker/modsprofile` and sheet `tiff_images` to see validation messages. This workbook is intentionally invalid and should not download a ZIP when `Enforce validations` is checked. Uncheck `Enforce validations` to see warning-only processing for Preview and Download.

Expected row behavior:

- Row 2: missing `imageAccessibilityAltText` with `typeOfResource` set to `still image`; should fail required alt-text validation.
- Row 3: missing `typeOfResource`; should pass the conditional required rule.
- Row 4: `imageAccessibilityAltText` is longer than 250 characters; should fail max-length validation.
- Row 5: `typeOfResource` is ` Still Image ` with extra whitespace and different case; should fail required alt-text validation.
- Row 6: `typeOfResource` is `still image|text`; should pass because the rule only matches exactly `still image`.
- Row 7: normal still-image row with valid alt text; should pass.
- Row 8: normal non-image row with blank alt text; should pass.
