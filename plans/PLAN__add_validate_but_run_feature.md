# Plan: Allow Validation Warnings Without Stopping MODS Processing

## Overview

Add an upload-page option that lets users choose whether MODS validation failures stop processing. The default should remain strict: validation failures block Preview and Download. If the user disables enforcement, Preview and Download should still run, but the page should show validation warnings and explain that generated MODS with failed validation may not work in the Workshop.

Recommended checkbox label:

```text
Enforce validations
```

Default state: checked.

## Contents

- [Overview](#overview)
- [Current Behavior](#current-behavior)
- [Desired Behavior](#desired-behavior)
- [Implementation Approach](#implementation-approach)
- [Backend Changes](#backend-changes)
- [Template Changes](#template-changes)
- [Warning Display](#warning-display)
- [Tests](#tests)
- [Documentation Updates](#documentation-updates)
- [Resolved Decisions](#resolved-decisions)
- [Implementation Order](#implementation-order)

## Current Behavior

- Validation rules live in active MODS YAML profiles under `validations:`.
- Validation logic lives in `profileValidation.py`.
- `fileSupport.validateRowsForProfile()` raises `profileValidation.ValidationError` when any validation rule fails.
- `fileSupport.createZipFromExcel()` validates before generating a ZIP.
- `fileSupport.getPreview()` validates before generating preview text.
- `/modsmaker/getpreview` returns JSON validation errors instead of preview text when validation fails.
- `POST /modsmaker/<profileFilename>` renders the generic error page when validation fails.
- The current upload-page JavaScript is still relatively small; preserve that. New behavior should mostly live in Flask/Python helper code.

## Desired Behavior

When `Enforce validations` is checked:

- Keep current behavior.
- Preview stops and shows validation errors.
- Download stops and shows validation errors.
- This remains the default.

When `Enforce validations` is unchecked:

- Preview still generates MODS preview output.
- Download shows validation warnings before ZIP generation, then still allows ZIP generation.
- Validation failures are shown as warnings, not blocking errors.
- Warning text should be visible near the preview/output workflow and should include wording like:

```text
Validation warnings were found. Processing continued because validations are not being enforced. MODS created with failed validation may not work in the Workshop.
```

Important behavior:

- Validation should still run in both modes.
- The checkbox changes whether validation failures are blocking.
- The checkbox should not change the actual YAML validation rules.

## Implementation Approach

Prefer a server-side implementation with minimal JavaScript.

Add a request flag named something like:

```text
enforce_validations
```

Default form behavior:

- The checkbox is checked by default.
- When checked, the form includes `enforce_validations`.
- When unchecked, the form omits it.

For async preview requests:

- Existing JavaScript already sends a JSON `data` payload.
- Include `enforce_validations: true/false` in that payload.
- Keep JS responsibility limited to reading the checkbox and displaying the returned response.

For Download:

- Add a lightweight server-side validation/preflight endpoint used before form submission.
- The endpoint should return server-formatted error/warning text.
- JavaScript should only call that endpoint, render the returned text, and either stop or continue the normal form submission.
- Keep validation rules, message formatting, and enforcement decisions in Flask/Python.

Backend behavior:

- Convert the flag to a boolean in Flask.
- Pass it to `fileSupport.getPreview()` and `fileSupport.createZipFromExcel()`.
- Let `fileSupport` decide whether validation errors raise or become warnings.

## Backend Changes

Update helper signatures:

```python
createZipFromExcel(excelFile, sheetName, profilePath, globalConditions, enforceValidations=True)
getPreview(excelFile, sheetName, profilePath, globalConditions, enforceValidations=True)
```

Add helper behavior:

```python
def getValidationWarningsForProfile(rows, profilePath):
    ...
```

or adjust `validateRowsForProfile()` to optionally return errors instead of raising:

```python
def validateRowsForProfile(rows, profilePath, raiseOnError=True):
    ...
```

Recommended shape:

- Keep `validateRowsForProfile(rows, profilePath)` as the strict helper to avoid changing existing call semantics accidentally.
- Add a new helper such as `getValidationErrorsForProfile(rows, profilePath)` that returns errors and respects skipped rows.
- Have strict flows raise if returned errors exist.
- Have warning flows continue generation and return warnings to the route.

Preview route response when enforcement is disabled and warnings exist:

```json
{
  "preview": "...generated MODS preview...",
  "warnings": [...],
  "warning_text": "Row 2, column \"imageAccessibilityAltText\": ..."
}
```

Preview route response when enforcement is enabled and errors exist:

```json
{
  "errors": [...],
  "error_text": "Row 2, column \"imageAccessibilityAltText\": ..."
}
```

Download behavior when enforcement is disabled and warnings exist:

- Direct Download must show warnings before the ZIP download starts.
- Add a lightweight `/modsmaker/validate` endpoint for preflight validation.
- Keep this endpoint Python-owned:
  - parse workbook rows,
  - run YAML validation,
  - decide whether failures are blocking errors or non-blocking warnings,
  - format display text server-side.
- Keep JavaScript minimal:
  - read the `Enforce validations` checkbox,
  - call `/modsmaker/validate`,
  - render returned `error_text` or `warning_text`,
  - submit the existing form only when processing may continue.

## Template Changes

Add a checkbox near the sheet selector/global conditions area:

```html
<div class="form-check">
  <input class="form-check-input" type="checkbox" name="enforce_validations" id="enforce_validations" checked>
  <label class="form-check-label" for="enforce_validations">Enforce validations</label>
</div>
```

Possible helper text:

```text
When unchecked, Preview and Download continue even if validation warnings are found.
```

Prefer the label `Enforce validations`. It is concise and matches the system behavior.

## Warning Display

For Preview:

- If warnings exist and preview still runs, display warning text above the generated MODS preview.
- Suggested display text:

```text
Validation warnings were found. Processing continued because validations are not being enforced. MODS created with failed validation may not work in the Workshop.

Row 2, column "imageAccessibilityAltText": ...
```

Implementation options:

- Best small option: have Flask return one combined `warning_text` string and have JS prepend it to preview text.
- Alternative: add a second warning area in `templates/preview.html`.

For Download:

- Direct Download must show warning text before ZIP generation when enforcement is off and validation warnings exist.
- Use the same display area and server-formatted warning text as Preview.
- After warning text is displayed, continue normal form submission automatically if validations are not enforced.
- When enforcement is on, validation failures should display as errors and stop form submission.

## Tests

Add focused tests for helper and route behavior.

Backend helper tests:

- Strict mode with invalid still-image row raises `ValidationError`.
- Non-enforcing mode with invalid still-image row returns warnings and still creates preview text.
- Non-enforcing mode with invalid still-image row still creates a ZIP.

Flask route tests:

- `/modsmaker/getpreview` with invalid spreadsheet and `enforce_validations: true` returns `errors`.
- `/modsmaker/getpreview` with invalid spreadsheet and `enforce_validations: false` returns preview text plus `warnings` or `warning_text`.
- `/modsmaker/validate` with invalid spreadsheet and `enforce_validations: true` returns blocking `errors` and `error_text`.
- `/modsmaker/validate` with invalid spreadsheet and `enforce_validations: false` returns non-blocking `warnings`, `warning_text`, and a continue/process flag.
- `POST /modsmaker/<profile>` with invalid spreadsheet and checked checkbox returns validation error page.
- `POST /modsmaker/<profile>` with invalid spreadsheet and unchecked checkbox returns a ZIP.

Existing tests to update:

- Existing preview validation tests should specify the enforcing/default behavior.
- Existing invalid download tests should specify checked enforcement.
- Add tests using the validation-example workbook if useful.

## Documentation Updates

Update:

- `README.md`: mention the `Enforce validations` option in the MODS Maker section.
- `spreadsheet_demos/README.md`: mention that the validation example workbook can demonstrate both strict validation and warning-only processing.

Suggested language:

```text
By default, MODS validation errors stop Preview and Download. To inspect generated MODS despite validation failures, uncheck Enforce validations. The app will continue processing and show validation warnings, but generated MODS may not work in the Workshop.
```

## Resolved Decisions

No policy decisions remain open.

Resolved behavior:

1. Download warning display

- Direct Download must show validation warnings before starting ZIP generation when validation enforcement is disabled.
- Use a lightweight `/modsmaker/validate` endpoint for this preflight.
- Keep JavaScript minimal by having Flask return server-formatted warning text.

2. Wording

Use this checkbox label:

```text
Enforce validations
```

Use this warning sentence:

```text
Processing continued because validations are not being enforced. MODS created with failed validation may not work in the Workshop.
```

3. JavaScript scope

- JavaScript may coordinate browser-only flow: collect the selected file, call `/modsmaker/validate`, render server-formatted text, and submit the form.
- Validation rules, error/warning classification, and message formatting should remain in Flask/Python.

4. Response shape

- For minimal disruption, keep existing successful preview string response in strict mode if practical.
- Return an object when warnings are present.
- `/modsmaker/validate` should always return a consistent object.
- A future cleanup can normalize all preview responses.

## Implementation Order

1. Add checkbox to `templates/mods/modsAction.html` or adjacent upload controls, default checked.
2. Add a small `/modsmaker/validate` endpoint that returns server-formatted errors or warnings.
3. Add minimal JS to include `enforce_validations` in preview and validation request data.
4. Parse `enforce_validations` in `modsMakerGetPreview()`.
5. Parse `enforce_validations` in `modsMakerValidate()`.
6. Parse `enforce_validations` in `modsMakerHome()` from normal form data.
7. Add non-raising validation helper in `fileSupport.py`.
8. Update `getPreview()` and `createZipFromExcel()` to support strict and warning-only modes.
9. Update Preview flow so warning text appears before generated MODS when enforcement is off.
10. Update Download flow so it calls `/modsmaker/validate`, displays server-formatted warnings, and then submits the form when enforcement is off.
11. Add route/helper tests for strict versus warning-only modes.
12. Update README and spreadsheet demo documentation.
13. Run `uv run ./run_tests.py`.
