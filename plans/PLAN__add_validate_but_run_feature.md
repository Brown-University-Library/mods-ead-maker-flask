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
- [Open Decisions](#open-decisions)
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
- Download still generates the ZIP.
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

- A normal file download response cannot also display warnings in the current same-page UI.
- To keep JavaScript minimal, add a preview/validate preflight for Download only if the current UI already relies on async validation.
- If the working tree does not have a `/modsmaker/validate` endpoint, avoid adding a large JS workflow just for this feature. Instead, use a simple server-side approach first:
  - checked: current error page on validation failure.
  - unchecked: download ZIP even with warnings.
- If visible download warnings are required before download starts, add a lightweight `/modsmaker/validate` endpoint and minimal JS in a later pass.

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

- If avoiding a larger JS flow, unchecked enforcement should allow the ZIP to download even if warnings exist.
- The warning state will be visible in Preview but not during a direct Download click.
- If direct Download warning visibility is important, add a small preflight endpoint and display warnings before submitting.

## Tests

Add focused tests for helper and route behavior.

Backend helper tests:

- Strict mode with invalid still-image row raises `ValidationError`.
- Non-enforcing mode with invalid still-image row returns warnings and still creates preview text.
- Non-enforcing mode with invalid still-image row still creates a ZIP.

Flask route tests:

- `/modsmaker/getpreview` with invalid spreadsheet and `enforce_validations: true` returns `errors`.
- `/modsmaker/getpreview` with invalid spreadsheet and `enforce_validations: false` returns preview text plus `warnings` or `warning_text`.
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

## Open Decisions

1. Download warning display

Decide whether unchecked Download must show warnings before starting the ZIP download.

Recommendation for smallest implementation:

- Do not add a larger JavaScript preflight flow yet.
- Let unchecked Download create the ZIP.
- Rely on Preview for visible warning review.

Recommendation for best UX:

- Add or restore a lightweight `/modsmaker/validate` endpoint so Download can show warnings before submitting.
- Keep JavaScript minimal by having Flask return server-formatted warning text.

2. Wording

Recommended checkbox label:

```text
Enforce validations
```

Recommended warning sentence:

```text
Processing continued because validations are not being enforced. MODS created with failed validation may not work in the Workshop.
```

3. Response shape

Decide whether preview responses should remain string-or-error-object mixed, or move to a consistent object response.

Recommendation:

- For minimal disruption, keep existing successful preview string response in strict mode.
- Return an object only when warnings are present.
- A future cleanup can normalize preview responses.

## Implementation Order

1. Add checkbox to `templates/mods/modsAction.html` or adjacent upload controls, default checked.
2. Add minimal JS to include `enforce_validations` in preview request data.
3. Parse `enforce_validations` in `modsMakerGetPreview()`.
4. Parse `enforce_validations` in `modsMakerHome()` from normal form data.
5. Add non-raising validation helper in `fileSupport.py`.
6. Update `getPreview()` and `createZipFromExcel()` to support strict and warning-only modes.
7. Update preview rendering so warning text appears before generated MODS when enforcement is off.
8. Add route/helper tests for strict versus warning-only modes.
9. Update README and spreadsheet demo documentation.
10. Run `uv run ./run_tests.py`.
