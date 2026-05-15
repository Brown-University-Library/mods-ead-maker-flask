# Plan: Add MODS Image Accessibility Alt Text

## Overview

The MODS Maker should be able to generate an image accessibility note from a spreadsheet column:

```xml
<mods:note type="image_accessibility_alt_text">The image accessibility text.</mods:note>
```

The existing profile interpreter already supports creating `mods:note` elements from YAML profile fields, so the smallest implementation is profile-driven: add a new column mapping to the relevant MODS YAML profile or profiles. The main open question is how strongly the 250-character workshop rule should be enforced in the webapp, because the current app has no general profile-validation or row-error reporting layer.

## Contents

- [Overview](#overview)
- [Current Behavior](#current-behavior)
- [Recommended Implementation](#recommended-implementation)
- [YAML Profile Change](#yaml-profile-change)
- [Spreadsheet Column](#spreadsheet-column)
- [Character Limit Handling](#character-limit-handling)
- [Tests](#tests)
- [Documentation Updates](#documentation-updates)
- [Decision Points](#decision-points)
- [Implementation Order](#implementation-order)
- [Original Prompt](#original-prompt)

## Current Behavior

- MODS profile selection is route-driven. For example, `/modsmaker/modsprofile` uses `profiles/modsprofile.yaml`.
- The uploaded spreadsheet does not choose the profile.
- `fileSupport.py` reads rows from the selected `.xlsx` sheet into dictionaries keyed by spreadsheet header.
- `profileInterpreter.Profile` reads the active YAML profile and turns each row into MODS XML.
- Existing profiles already create many `mods:note` elements using YAML blocks like:

```yaml
- type: element
  name: note
  attrs: {type: general}
  text:
    - type: value
      values:
        - {type: col, header: noteGeneral, method: value}
```

- Empty generated elements are removed during XML cleanup, so a blank alt-text column should not produce an empty `mods:note`.
- There is no current app-level validation system for maximum field length, required fields, row warnings, or blocking upload errors.

## Recommended Implementation

Add the alt-text note as a normal YAML-driven MODS field first. This keeps the mapping transparent, profile-specific, and consistent with the rest of the app.

Recommended output:

```xml
<mods:note type="image_accessibility_alt_text">A concise description of the image.</mods:note>
```

Recommended spreadsheet column:

```text
imageAccessibilityAltText
```

This column name is explicit enough to distinguish front-end alt-text display data from general descriptive notes.

## YAML Profile Change

Add this field to each profile that should support image accessibility alt text:

```yaml
- type: element
  name: note
  attrs: {type: image_accessibility_alt_text}
  text:
    - type: value
      values:
        - {type: col, header: imageAccessibilityAltText, method: value}
```

Recommended first target:

- `profiles/modsprofile.yaml`

Potential additional targets:

- `profiles/hallhoag.yaml`, if Hall-Hoag image MODS will use the same workflow.
- `profiles/jnbcsyllabi.yaml`, `profiles/musictheses.yaml`, and `profiles/musicdoctoraldissertation.yaml` only if those profiles may describe image resources or the workshop expects a consistent column across all MODS profiles.
- Avoid editing `profiles/modsprofile_backup2024.yaml` unless the backup file is intentionally maintained as a live profile.

Placement recommendation:

- Put the new field near the other `note` fields in each profile.
- Do not make it conditional on `typeOfResource` in the first pass unless the team wants non-image rows to reject or ignore the column. A blank column will naturally produce no note.

## Spreadsheet Column

Add `imageAccessibilityAltText` to relevant templates and demo spreadsheets.

For image workflows, the workshop should document:

- The field is required for images.
- The maximum value is 250 characters.
- The value should be written as display-ready alt text, not as a technical filename, identifier, or generic phrase.

Recommended demo update:

- Add the column to `spreadsheet_demos/mods_default_tiff_images_basic.xlsx`.
- Add values under 250 characters.
- Update `spreadsheet_demos/README.md` to mention that the TIFF image demo includes the accessibility note field.

## Character Limit Handling

The 250-character rule can be handled at three possible levels.

Option A: Documentation-only first pass

- Add the YAML mapping and document the limit.
- Keep current generation behavior unchanged.
- Fastest and lowest-risk option.
- Risk: records over 250 characters will still generate XML unless checked outside the app.

Option B: Test/demo enforcement only

- Add tests that confirm demo values stay at or below 250 characters.
- Keep app behavior unchanged for arbitrary user uploads.
- Useful for workshop materials, but not true user-input enforcement.

Option C: App-level validation

- Add validation before preview/download that checks configured profile constraints.
- For `imageAccessibilityAltText`, block or warn when a non-empty value exceeds 250 characters.
- This requires designing how row-level errors are returned in preview and download flows.
- Best long-term enforcement, but larger scope than a simple profile mapping.

Recommended path:

1. Implement Option A plus tests that confirm the XML output.
2. Decide whether the workshop needs hard blocking before launch.
3. If hard blocking is needed, implement a small profile-aware validation layer rather than hard-coding this one column deeply inside XML generation.

## Tests

Add focused tests around the default MODS profile.

Recommended tests:

- `profileInterpreter.Profile('profiles/modsprofile.yaml').convertRowToXmlString(...)` creates a `mods:note` with `type="image_accessibility_alt_text"` when `imageAccessibilityAltText` is present.
- The generated XML does not contain that note when `imageAccessibilityAltText` is blank or absent.
- `fileSupport.createZipFromExcel()` preserves the note when processing an uploaded workbook.
- If app-level validation is implemented, add route tests for preview/download responses when the value exceeds 250 characters.

For the demo spreadsheet:

- Extend `tests/test_spreadsheet_demos.py` so the TIFF image demo verifies the generated XML includes the new note.
- Optionally assert all demo `imageAccessibilityAltText` values are `<= 250` characters.

## Documentation Updates

Update:

- Main `README.md`: mention the alt-text column in the MODS Maker overview or spreadsheet guidance.
- `spreadsheet_demos/README.md`: mention which demo includes `imageAccessibilityAltText`.
- Any workshop instructions or external spreadsheet templates.

Suggested documentation language:

```text
For image records, include an imageAccessibilityAltText column. When present, the MODS Maker writes it as <mods:note type="image_accessibility_alt_text">...</mods:note>. Workshop image records should keep this value at or below 250 characters.
```

## Decision Points

1. Column name

Recommended: `imageAccessibilityAltText`.

Alternatives:

- `imageAltText`
- `altText`
- `noteImageAccessibilityAltText`

2. Profile scope

Decide whether to add the field only to `modsprofile.yaml` or to every active MODS profile.

3. Enforcement level

Decide whether the 250-character limit is:

- documented only,
- tested only for demos/templates,
- a warning in preview,
- or a blocking error for preview/download.

4. Requiredness

Decide whether "required for images" is enforced by:

- workshop/template review,
- app validation when `typeOfResource` indicates an image,
- or downstream indexing/QA.

5. Image detection

If requiredness is app-enforced, define exactly what counts as an image row. Possible signals:

- `typeOfResource` equals `still image`,
- TIFF/image-specific workflow spreadsheet,
- route/profile selection,
- a new explicit spreadsheet column.

6. Over-limit behavior

If a value is longer than 250 characters, decide whether to:

- block generation,
- show a warning but generate XML,
- truncate automatically,
- or leave the value unchanged.

Recommendation: do not truncate automatically. It can silently change cataloging intent.

7. Backup profile handling

Decide whether `modsprofile_backup2024.yaml` should stay historical or receive the same mapping.

## Implementation Order

1. Confirm the column name and target profile list.
2. Add the YAML field to `profiles/modsprofile.yaml`.
3. Add or update tests proving the new note appears in generated MODS XML.
4. Update the TIFF demo spreadsheet with `imageAccessibilityAltText` values.
5. Update demo tests and README documentation.
6. Decide whether to add validation for the 250-character limit.
7. If validation is needed, add a small profile-aware validation mechanism and route tests.

## Original Prompt

```text
Goal: add mods-alt-text note.

Context:

- Due to accessibility regulations (and good practice), we're beginning to enforce that images have image-description-text that we'll use for front-end alt-text displays.

- A decision was made to implement it like this, for MODS:

"""
<mods:note type="image_accessibility_alt_text">The image accessibility text.</mods:note>
"""

- The workshop will require this for images with MODS, with a maximum character-limit of 250 characters.

- Indexing infrastructure has already been implemented.

Tasks:

- Come up with a plan to update the mods-maker code and/or yaml-files to create mods with this `mods:note` element -- given a specified spreadsheet column.

- Save the plan to `mods-ead-maker-flask/plans/PLAN__add_mods_a11y_data.md`.

- Note any decision-points that may need to be addressed.

- Append this prompt to the bottom of the document.

- From this prompt, and the actual plan -- add a brief overview section near the top.

- Also add, near the top, a table-of-contents/index with links to the headings in the plan.

Thx!
```
