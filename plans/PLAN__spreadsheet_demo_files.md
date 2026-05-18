# Spreadsheet Demo Files Plan

Goal: add a small set of spreadsheet files that let users quickly see the MODS Maker workflow in action.

Status:

- Done: created the recommended MODS demo spreadsheets in `spreadsheet_demos/`.
- Done: added `mods_default_tiff_images_basic.xlsx` for still-image/TIFF-oriented MODS experiments.
- Done: created `spreadsheet_demos/README.md`.
- Done: validated each spreadsheet through `fileSupport.createZipFromExcel()` and confirmed generated MODS XML parses.
- Next: add EAD demo spreadsheets later, after a minimal EAD workbook is confirmed.


## Recommendation

Start with MODS demo spreadsheets only.

Reasons:

- The MODS Maker has the clearest demo path: open `/modsmaker/<profile>`, upload `.xlsx`, choose a sheet, preview records, and download a ZIP of generated `.mods.xml` files.
- The MODS profiles are visible in the UI and have clear YAML-driven behavior worth demonstrating.
- EAD generation is user-facing, but it depends more on legacy code and the `legacy/cache` upload flow. EAD demo files can be added later once the MODS demos are stable and documented.
- A small MODS-only first set avoids overwhelming users with too many spreadsheet shapes.

Suggested follow-up: add EAD examples later under a separate `spreadsheet_demos/ead/` folder after confirming a minimal EAD workbook that exercises the legacy EAD flow safely.


## Proposed Directory Layout

```text
spreadsheet_demos/
  README.md
  mods_default_basic.xlsx
  mods_default_repeating_fields.xlsx
  mods_default_tiff_images_basic.xlsx
  mods_john_nicholas_brown_center_syllabi_basic.xlsx
  mods_music_theses_basic.xlsx
```

Optional later layout:

```text
spreadsheet_demos/
  mods/
    ...
  ead/
    ...
```

For the first pass, keeping the files directly in `spreadsheet_demos/` is simpler for users.


## Demo 1: `mods_default_basic.xlsx`

Profile to use:

- `/modsmaker/modsprofile`

Purpose:

- Demonstrates the default MODS profile with a small, approachable metadata record.
- Shows title, subtitle, dates, language, type/resource, identifiers, repository/context fields, and rights/access notes.
- Gives a user a quick success case that produces one or two `.mods.xml` files.

Suggested sheet name:

- `basic_records`

Suggested row count:

- 2 records

Suggested columns:

- `identifierFileName`
- `fileTitle`
- `itemTitle`
- `subTitle`
- `dateText`
- `dateStart`
- `dateEnd`
- `typeOfResource`
- `language`
- `abstract`
- `noteGeneral`
- `identifierBDR`
- `collection`
- `repository`
- `findingAid`
- `rightsStatementText`
- `rightsStatementURI`
- `useAndReproduction`

What it demonstrates:

- Basic row-to-MODS XML conversion.
- Filename generation from `identifierFileName`.
- MODS title generation.
- Common descriptive and administrative metadata fields.
- Global condition checkboxes for Brown defaults and preferred citation.


## Demo 2: `mods_default_repeating_fields.xlsx`

Profile to use:

- `/modsmaker/modsprofile`

Purpose:

- Demonstrates the more domain-specific behavior users are likely to need: repeated names, subjects, URIs, and multi-value cells.

Suggested sheet name:

- `repeating_fields`

Suggested row count:

- 2 records

Suggested columns:

- `identifierFileName`
- `fileTitle`
- `namePersonCreatorLC`
- `namePersonCreatorLocal`
- `nameCorpCreatorLocal`
- `namePersonOtherLocal`
- `subjectTopicsLC`
- `subjectTopicsLocal`
- `subjectGeoLC`
- `subjectTemporalLC`
- `subjectNamesLocal`
- `genreAAT`
- `language`
- `identifierLocal`

Suggested data patterns:

- Use semicolon-separated or pipe-separated repeated values where the profile supports repeating values.
- Include a creator with a date, such as `Doe, Jane, 1970-`.
- Include an "other" name with a role, such as `Smith, Alex, 1980-, Photographer`.
- Include a URI in at least one authority-controlled value.
- Include one bracketed additional-value example if appropriate, such as `[displayForm: Jane Doe]`.

What it demonstrates:

- Repeating name fields.
- Repeating subject fields.
- URI extraction into authority/valueURI attributes.
- Name/date/role parsing.
- Local versus authority-backed values.


## Demo 3: `mods_default_tiff_images_basic.xlsx`

Profile to use:

- `/modsmaker/modsprofile`

Purpose:

- Demonstrates default-profile MODS records shaped around TIFF/still-image metadata.
- Gives users a concrete starting point for experimenting with MODS records that conceptually pair with source image files.

Suggested sheet name:

- `tiff_images`

Suggested row count:

- 2 records

Suggested columns:

- `identifierFileName`
- `fileTitle`
- `itemTitle`
- `dateText`
- `dateStart`
- `dateEnd`
- `typeOfResource`
- `genreAAT`
- `digitalOrigin`
- `form`
- `extentQuantity`
- `extentSize`
- `language`
- `abstract`
- `noteGeneral`
- `subjectTopicsLocal`
- `subjectGeoLC`
- `identifierBDR`
- `identifierLocal`
- `collection`
- `repository`
- `findingAid`
- `rightsStatementText`
- `rightsStatementURI`
- `useAndReproduction`

What it demonstrates:

- `identifierFileName` values that align with source TIFF basenames, such as `demo_image_0001`.
- Still-image resource metadata.
- TIFF-oriented form and extent fields.
- Image rights, repository, collection, and local identifier fields.


## Demo 4: `mods_john_nicholas_brown_center_syllabi_basic.xlsx`

Profile to use:

- `/modsmaker/jnbcsyllabi`

Purpose:

- Demonstrates a profile-specific workflow that differs from the default archival/item MODS profile.
- Useful for showing that the app is not hard-coded to one spreadsheet shape.

Suggested sheet name:

- `syllabi`

Suggested row count:

- 2 records

Suggested columns:

- `fileno`
- `courseTitle`
- `subTitle`
- `Abstract`
- `instructor (invert name)`
- `date`
- `description`
- `courseno`
- `language`

What it demonstrates:

- Profile-specific filename column: `fileno`.
- Course title to MODS title generation.
- Instructor name as creator.
- Course number and description metadata.


## Demo 5: `mods_music_theses_basic.xlsx`

Profile to use:

- `/modsmaker/musictheses`

Purpose:

- Demonstrates a specialized profile for thesis/dissertation-style records.
- Useful for users who need to see a non-archival, publication-like MODS record.

Suggested sheet name:

- `music_theses`

Suggested row count:

- 2 records

Suggested columns:

- `identifierBDR`
- `Project title`
- `subTitle`
- `Abstract`
- `Author name`
- `Advisor name`
- `Date created`
- `Date`
- `Type of resource`
- `Concentration`
- `language`
- `License preference`
- `License URI`
- `License name`
- `collection`
- `repository`
- `findingAid`

What it demonstrates:

- Profile-specific filename column: `identifierBDR`.
- Thesis-style title, author, advisor, date, concentration, and license metadata.
- `keepblanktextelements` behavior for license/logo-related profile fields.


## `spreadsheet_demos/README.md` Recommendation

Add a README next to the spreadsheets with:

- The exact route/profile to use for each demo spreadsheet.
- The sheet name to select after upload.
- A short explanation of what each file demonstrates.
- A note that generated output is downloaded as a ZIP of `.mods.xml` files.
- A note that demo metadata is fictional or sample-only.
- A reminder that EAD examples are planned for a later pass.

Suggested README table:

```markdown
| File | Route | Sheet | Demonstrates |
| --- | --- | --- | --- |
| `mods_default_basic.xlsx` | `/modsmaker/modsprofile` | `basic_records` | Basic MODS fields and download flow |
| `mods_default_repeating_fields.xlsx` | `/modsmaker/modsprofile` | `repeating_fields` | Names, subjects, repeated values, URIs |
| `mods_default_tiff_images_basic.xlsx` | `/modsmaker/modsprofile` | `tiff_images` | Still-image/TIFF-oriented MODS fields |
| `mods_john_nicholas_brown_center_syllabi_basic.xlsx` | `/modsmaker/jnbcsyllabi` | `syllabi` | John Nicholas Brown Center syllabi profile |
| `mods_music_theses_basic.xlsx` | `/modsmaker/musictheses` | `music_theses` | Music thesis metadata profile |
```


## Implementation Notes

- Use `.xlsx`, not `.csv`, because the upload flow checks for `.xlsx`.
- Keep files small: two rows per demo is enough.
- Use realistic but fictional data.
- Prefer field names that appear directly in the YAML profiles.
- Include all required filename columns so generated output filenames are clear.
- Include at least one intentionally sparse row in each spreadsheet so users can see that blank fields are removed from generated XML.
- Avoid examples that depend on external network access.
- Avoid very large profiles or edge-case-only fields in the first pass.


## Validation Steps

For each demo spreadsheet:

1. Start the app.
2. Open the listed `/modsmaker/<profile>` route.
3. Upload the spreadsheet.
4. Select the listed sheet.
5. Preview the generated XML.
6. Download the ZIP.
7. Confirm the ZIP contains one `.mods.xml` file per non-skipped row.
8. Confirm each XML file opens and contains the expected title/identifier/name fields.


## Later EAD Demo Plan

Add EAD examples only after confirming:

- the minimum workbook columns needed for a useful EAD output,
- whether the existing `legacy/Collection-Level Data.xlsx` can safely serve as a model,
- whether any generated cache files need cleanup guidance,
- and how to explain EAD output expectations clearly to non-developer users.

Possible later file:

- `ead_collection_basic.xlsx`

For now, keep EAD out of the first demo batch.
