# Plan: Add MODS Image Accessibility Alt Text

## Overview

The MODS Maker should be able to generate an image accessibility note from a spreadsheet column:

```xml
<mods:note type="image_accessibility_alt_text">The image accessibility text.</mods:note>
```

The existing profile interpreter already supports creating `mods:note` elements from YAML profile fields, so the MODS output mapping can be profile-driven. The 250-character rule should also be profile-driven by adding a generic `validations:` section to the active YAML profile. Python should implement the validation engine once, while each profile declares which spreadsheet columns have limits or conditional requirements.

## Contents

- [Overview](#overview)
- [Current Behavior](#current-behavior)
- [Recommended Implementation](#recommended-implementation)
- [YAML Profile Change](#yaml-profile-change)
- [Spreadsheet Column](#spreadsheet-column)
- [YAML Validation Design](#yaml-validation-design)
- [Validation Behavior](#validation-behavior)
- [Feedback Incorporated](#feedback-incorporated)
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

Add the alt-text note as a normal YAML-driven MODS field, and add a profile-level `validations:` section that enforces the 250-character limit before preview or download output is produced.

This keeps two responsibilities separate:

- `fields:` controls generated MODS XML.
- `validations:` controls spreadsheet-row rules.

Recommended output:

```xml
<mods:note type="image_accessibility_alt_text">A concise description of the image.</mods:note>
```

Recommended spreadsheet column:

```text
imageAccessibilityAltText
```

This column name is explicit enough to distinguish front-end alt-text display data from general descriptive notes.

Recommended validation approach:

```yaml
validations:
  - type: maxchars
    col: imageAccessibilityAltText
    maxchars: 250
    severity: error
    message: "Image accessibility alt text must be 250 characters or fewer."
  - type: required
    col: imageAccessibilityAltText
    severity: error
    conditions:
      - {type: equals, col: typeOfResource, text: "still image"}
    message: "Image accessibility alt text is required when typeOfResource is still image."
```

## YAML Profile Change

Add this output field to each profile that should support image accessibility alt text:

```yaml
- type: element
  name: note
  attrs: {type: image_accessibility_alt_text}
  text:
    - type: value
      values:
        - {type: col, header: imageAccessibilityAltText, method: value}
```

Recommended target scope:

- every active MODS profile

Likely active profile files:

- `profiles/modsprofile.yaml`
- `profiles/hallhoag.yaml`
- `profiles/jnbcsyllabi.yaml`
- `profiles/musictheses.yaml`
- `profiles/musicdoctoraldissertation.yaml`

Explicitly excluded:

- `profiles/modsprofile_backup2024.yaml`

Placement recommendation:

- Put the new field near the other `note` fields in each profile.
- Keep the output field itself unconditional. A blank or missing `imageAccessibilityAltText` column will naturally produce no note after XML cleanup. Requiredness belongs in `validations:`, not in the output mapping.

Add this validation block near the existing top-level profile settings, alongside keys such as `globalconditions`, `filenamecolumn`, and `fileextension`:

```yaml
validations:
  - type: maxchars
    col: imageAccessibilityAltText
    maxchars: 250
    severity: error
    message: "Image accessibility alt text must be 250 characters or fewer."
```

Enforce requiredness for image rows with a conditional required rule:

```yaml
validations:
  - type: maxchars
    col: imageAccessibilityAltText
    maxchars: 250
    severity: error
    message: "Image accessibility alt text must be 250 characters or fewer."
  - type: required
    col: imageAccessibilityAltText
    severity: error
    conditions:
      - {type: equals, col: typeOfResource, text: "still image"}
    message: "Image accessibility alt text is required when typeOfResource is still image."
```

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

## YAML Validation Design

Add a small generic validation engine that reads a top-level `validations:` list from the active YAML profile.

Recommended first supported rule:

```yaml
- type: maxchars
  col: imageAccessibilityAltText
  maxchars: 250
  severity: error
  message: "Image accessibility alt text must be 250 characters or fewer."
```

Second supported rule:

```yaml
- type: required
  col: imageAccessibilityAltText
  severity: error
  conditions:
    - {type: equals, col: typeOfResource, text: "still image"}
  message: "Image accessibility alt text is required when typeOfResource is still image."
```

Recommended validation result shape:

```python
{
    "row_index": 2,
    "spreadsheet_row": 3,
    "col": "imageAccessibilityAltText",
    "type": "maxchars",
    "severity": "error",
    "message": "Image accessibility alt text must be 250 characters or fewer.",
    "value": "...",
    "limit": 250,
}
```

Design notes:

- Count Python string characters with `len(value)` after converting spreadsheet values to strings.
- Trim surrounding whitespace before length checks unless the team explicitly wants pasted leading/trailing spaces counted.
- Trim surrounding whitespace before condition comparisons.
- Compare `equals` condition values case-insensitively.
- For the `typeOfResource` requirement, the normalized value must be exactly `still image`. Values such as `still image|text`, `still image; text`, `moving image`, or any other non-exact value should waive the required alt-text rule.
- If the `typeOfResource` column is missing or blank, the required alt-text condition is false.
- Do not truncate values automatically.
- Keep validation separate from XML generation so the same rules can run before preview and before ZIP download.
- Support `maxchars`, `required`, and `equals` conditions in the first implementation.

Potential implementation locations:

- Add `self.profileValidations = self.profile.get("validations", [])` to `profileInterpreter.Profile`.
- Add a method such as `validateRow(row, rowIndex)` or `validateRows(rows)`.
- Alternatively, create a small `profileValidation.py` helper if this starts to grow beyond a few rule types.

## Validation Behavior

Enforce validation before producing preview or download output.

Preview flow:

- `/modsmaker/getpreview` reads the workbook rows for the selected sheet.
- The app validates rows against the active profile.
- If validation errors exist, return structured JSON containing errors instead of generated XML.
- The front end displays the errors in the preview area or a nearby alert.

Download flow:

- `POST /modsmaker/<profileFilename>` reads the workbook rows for the selected sheet.
- The app validates rows against the active profile.
- If validation errors exist, render an error page or return a user-readable validation page instead of downloading the ZIP.
- If no errors exist, continue generating the ZIP.

Recommended first UI behavior:

- Treat `severity: error` as blocking for both preview and download.
- Include spreadsheet row numbers, column names, and messages.
- Avoid warning-only behavior until the error path is clear.

## Feedback Incorporated

Feedback decisions now reflected in this plan:

- Add the MODS alt-text output mapping broadly to active MODS profiles. If a spreadsheet does not include `imageAccessibilityAltText`, generation should continue normally and no empty note should remain in the cleaned XML.
- Do not apply this change to `profiles/modsprofile_backup2024.yaml`.
- Add `maxchars` validation broadly to active MODS profiles. Missing or blank `imageAccessibilityAltText` values pass `maxchars`.
- Require `imageAccessibilityAltText` only if the row's `typeOfResource` value is exactly `still image` after trimming whitespace and normalizing case.
- If `typeOfResource` is missing, blank, or anything other than exactly `still image`, waive the required alt-text rule.
- Do not treat multiple-value strings as matches for the required rule. For example, `still image|text` and `still image; text` are not exactly `still image`, so they do not trigger required alt text.

## Tests

Add focused tests around the default MODS profile.

Recommended tests:

- `profileInterpreter.Profile('profiles/modsprofile.yaml').convertRowToXmlString(...)` creates a `mods:note` with `type="image_accessibility_alt_text"` when `imageAccessibilityAltText` is present.
- The generated XML does not contain that note when `imageAccessibilityAltText` is blank or absent.
- `fileSupport.createZipFromExcel()` preserves the note when processing an uploaded workbook.
- The profile validation method returns no errors for values at or below 250 characters.
- The profile validation method returns a blocking error for values over 250 characters.
- The profile validation method requires `imageAccessibilityAltText` when `typeOfResource` is `still image`.
- The required validation trims whitespace and is case-insensitive for `typeOfResource`, so values like ` Still Image ` trigger the requirement.
- The required validation does not trigger when `typeOfResource` is missing, blank, `moving image`, or a multiple-value string such as `still image|text`.
- Preview route returns validation errors instead of XML when `imageAccessibilityAltText` exceeds 250 characters.
- Download route does not return a ZIP when `imageAccessibilityAltText` exceeds 250 characters.

For the demo spreadsheet:

- Extend `tests/test_spreadsheet_demos.py` so the TIFF image demo verifies the generated XML includes the new note.
- Assert all demo `imageAccessibilityAltText` values are `<= 250` characters.

## Documentation Updates

Update:

- Main `README.md`: mention the alt-text column in the MODS Maker overview or spreadsheet guidance.
- `spreadsheet_demos/README.md`: mention which demo includes `imageAccessibilityAltText`.
- Any workshop instructions or external spreadsheet templates.

Suggested documentation language:

```text
For image records, include an imageAccessibilityAltText column. When present, the MODS Maker writes it as <mods:note type="image_accessibility_alt_text">...</mods:note>. Values over 250 characters are blocked by profile validation. The column is required when typeOfResource is still image, using trimmed and case-insensitive comparison.
```

## Decision Points

No open policy decisions remain for the planned implementation.

Resolved decisions:

1. Column name

Use `imageAccessibilityAltText`.

2. Profile scope

Add the output mapping and validation rules to every active MODS profile. Do not modify `profiles/modsprofile_backup2024.yaml`.

3. Validation rule format

Use this top-level YAML shape for profile validation rules:

```yaml
validations:
  - type: maxchars
    col: imageAccessibilityAltText
    maxchars: 250
    severity: error
    message: "Image accessibility alt text must be 250 characters or fewer."
```

4. Requiredness

Enforce requiredness in the app when `typeOfResource` is exactly `still image` after trimming whitespace and normalizing case.

If `typeOfResource` is missing, blank, or any value other than exactly `still image`, `imageAccessibilityAltText` is not required.

5. Image detection

For this validation rule, an image row is only a row whose normalized `typeOfResource` value is exactly `still image`.

Do not treat multiple-value strings as matches. Values such as `still image|text` or `still image; text` waive the required alt-text rule because they are not exactly `still image`.

6. Over-limit behavior

Block preview and download when `severity: error` validation fails. Do not truncate automatically; it can silently change cataloging intent.

7. Backup profile handling

Do not apply the change to `profiles/modsprofile_backup2024.yaml`.

Implementation details still to choose during coding:

- Whether validation helper code lives directly on `profileInterpreter.Profile` or in a small `profileValidation.py` module.
- Exact front-end presentation for preview validation errors.
- Exact non-preview error page/template used when download validation fails.

## Implementation Order

1. Add the YAML field, `maxchars` validation, and conditional `required` validation to every active MODS profile except `profiles/modsprofile_backup2024.yaml`.
2. Add generic profile validation support for `maxchars`, `required`, and `equals` conditions.
3. Implement trimmed, case-insensitive `equals` condition matching.
4. Ensure the `typeOfResource` required condition only matches exactly `still image` after normalization.
5. Run validation before preview and ZIP download generation.
6. Add or update tests proving the new note appears in generated MODS XML.
7. Add validation tests for values at, below, and above 250 characters.
8. Add validation tests for `typeOfResource` requiredness, including missing, blank, case/whitespace variants, non-image values, and multiple-value strings.
9. Update the TIFF demo spreadsheet with `imageAccessibilityAltText` values.
10. Update demo tests and README documentation.

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
