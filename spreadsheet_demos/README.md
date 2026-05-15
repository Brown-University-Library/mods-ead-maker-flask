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
- The TIFF image demo uses `identifierFileName` values like `demo_image_0001`, intended to pair conceptually with source files such as `demo_image_0001.tif`.
- The sample metadata is fictional and intended only for demonstration.
- EAD demo spreadsheets are not included yet; they can be added later after a minimal EAD example is confirmed.
