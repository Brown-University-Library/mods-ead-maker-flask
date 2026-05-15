# Test Plan

Goal: add useful tests without refactoring production code.

Status:

- Done: ran `uv run ./run_tests.py`; 27 tests pass.
- Done: reviewed `AGENTS.md`; no new test dependencies are needed because mocking uses `unittest.mock`.
- Done: replaced the template `tests/test.py` with real test modules for file helpers, profile parsing, MODS XML generation, and Flask routes.
- Next: add deeper EAD smoke tests and more upload/ZIP edge cases when useful.

The most-used project path appears to be the MODS Maker workflow:

1. User opens `/modsmaker` or `/modsmaker/<profile>`.
2. User uploads an `.xlsx` file and selects a sheet.
3. The app reads spreadsheet rows.
4. A YAML profile converts row metadata into MODS XML.
5. The app previews XML or returns generated XML files in a ZIP.

The EAD Maker routes are also user-facing, but they depend more heavily on legacy code and filesystem cache behavior. They should be tested after the core MODS path has coverage.


## Current Test State

- `run_tests.py` exists and uses `unittest`.
- `tests/test.py` is currently a template-style test that imports `main`, but this project does not have `main.py`.
- First test work should replace the template test with real project tests.
- No production-code refactor is required for the test layers below.


## Test Strategy

Use `unittest` and Flask's built-in test client.

Keep early tests focused on stable behavior that is both important and easy to exercise:

- pure helper functions in `fileSupport.py`
- parsing helpers in `profileInterpreter.py`
- high-value XML generation using real YAML profiles
- Flask route behavior for common GET routes and selected POST routes

Use small in-memory workbooks for spreadsheet tests. `xlrd==1.2.0` can read `.xlsx` bytes, and `openpyxl` or `xlsxwriter` can create fixture workbooks in memory.

Use `unittest.mock.patch` for route tests that do not need to exercise the full XML pipeline. This avoids refactoring while keeping route tests focused on request/response behavior.


## Phase 1: Replace Template Test

Status: implemented and verified.

Create a real test module structure, for example:

- `tests/test_file_support.py`
- `tests/test_profile_interpreter.py`
- `tests/test_flask_routes.py`

Remove the `main` import from the current template test.

Expected command:

```sh
uv run ./run_tests.py
```


## Phase 2: File Helper Tests

Status: implemented and verified.

Target: `fileSupport.py`

High-value tests:

- `cleanStringForFilename()` removes invalid filename characters.
- `getFilenameFromRow()` returns a cleaned filename from the configured filename column.
- `getFilenameFromRow()` falls back to `default<index>` when the filename column is blank or missing.
- `convertXlsxToDictList()` converts workbook rows into dictionaries keyed by header row.
- `convertXlsxToDictList()` serializes repeated headers with `|`.
- `getSheetNamesFromXlsx()` returns sheet names from uploaded workbook bytes.

Why this matters:

- Filename generation affects every downloaded MODS file.
- Spreadsheet parsing is the first step in the main upload workflow.
- Repeated spreadsheet headers are domain-specific behavior and easy to regress.


## Phase 3: Profile Parsing Tests

Status: implemented and verified.

Target: `profileInterpreter.py`

High-value tests:

- `normalizeString()` removes line breaks, repeated whitespace, pipes, and selected markup wrappers.
- `getValueUri()` extracts HTTP/HTTPS URIs from entries.
- `getAdditionalValues()` parses bracketed YAML additions, such as `[displayForm: Example]`.
- `getNameDateRoleFromEntry()` handles `nameCreator`.
- `getNameDateRoleFromEntry()` handles `nameOther`.
- `getMetadataFromEntry()` returns expected `entry.value`, `entry.valueURI`, `entry.name`, `entry.date`, `entry.role`, and additional values.
- `getKeyValueFromEntry()` parses `key: value`.

Why this matters:

- These helpers feed repeating MODS fields.
- They are relatively deterministic and can be tested without Flask or filesystem setup.
- They cover recent-risk behavior around bracketed additional values and colon-separated key/value parsing.


## Phase 4: MODS XML Generation Tests

Status: implemented and verified.

Target: `profileInterpreter.Profile` and `fileSupport.createFileFromRow()`

Use the real `profiles/modsprofile.yaml` unless the test becomes too brittle. If needed, add a tiny test-only YAML profile under `tests/fixtures/`.

High-value tests:

- `Profile('profiles/modsprofile.yaml').convertRowToXmlString(row)` returns MODS XML with the expected root namespace.
- A row with `fileTitle` or `itemTitle` creates a MODS title element.
- A row with `identifierFileName` produces a filename ending in `.mods.xml` through `createFileFromRow()`.
- A row with a configured `skipif` column populated returns `None`.
- Global conditions can include or suppress conditionally-generated elements, using the existing profile behavior.
- `createPreviewFromRows()` includes generated filenames and XML for multiple rows.
- `createZipFromExcel()` returns ZIP bytes containing expected generated filenames.

Why this matters:

- This is the core production behavior.
- It tests real profile semantics without changing production code.
- It gives confidence that spreadsheet rows still become usable MODS files.


## Phase 5: Flask Route Tests

Status: implemented and verified.

Target: `flask_app.py`

Use `flask_app.app.test_client()`.

GET route tests:

- `GET /modsmaker` redirects to `/modsmaker/modsprofile`.
- `GET /modsmaker/modsprofile` returns HTTP 200.
- `GET /profiles/` returns HTTP 200 and includes known profile names.
- `GET /forms/` returns HTTP 200 and includes known profile names.
- `GET /resources` returns HTTP 200.
- Unknown route redirects to the default MODS Maker profile.

POST route tests with mocks:

- `POST /processfileupload` with an `.xlsx` file returns JSON containing filename and sheet names.
- `POST /modsmaker/getpreview` calls `fileSupport.getPreview()` and returns JSON preview text.
- `POST /modsmaker/modsprofile` with a non-`.xlsx` filename returns the error template.
- `POST /forms/profile/modsprofile/preview` returns JSON preview text.

Why this matters:

- These tests validate the public app surface.
- Mocking the deep generation calls keeps route tests stable without refactoring.
- Full generation remains covered separately in helper/integration tests.


## Phase 6: EAD Smoke Tests

Status: not implemented yet; recommended next test layer.

Target: EAD routes in `flask_app.py` and selected `legacy/EADMaker.py` helpers.

Add these after MODS coverage is stable.

Candidate tests:

- `GET /eadmaker` returns HTTP 200.
- `POST /eadmaker` with a non-`.xlsx` file returns the error template.
- `GET /eadmaker/renderead/<filename>/<id>` can be tested with `getSheetNames` patched.
- `POST /eadmaker/renderead/<filename>/<id>` can be tested with `processExceltoEAD` patched.
- `POST /eadmaker/getpreview` can be tested with `processExceltoEAD` patched.

Why this matters:

- EAD remains user-facing.
- Patching avoids depending on legacy cache files or full workbook generation in route tests.


## Fixture Guidance

Prefer small test fixtures built inside tests:

- Build `.xlsx` bytes in memory.
- Use two sheets when testing sheet selection.
- Use repeated headers for repeated-column parsing.
- Keep rows minimal: one or two records per workbook.

For XML assertions:

- Prefer parsing output with `lxml.etree.fromstring()`.
- Assert on XPath results instead of long string comparisons.
- Use namespace-aware XPath for MODS.

For ZIP assertions:

- Use `zipfile.ZipFile(io.BytesIO(zip_bytes))`.
- Assert expected filenames exist.
- Assert file contents contain parseable XML.


## Risks And Notes

- Some existing production functions use relative paths internally. Tests should run from the project root, matching `run_tests.py` expectations.
- `tests/__pycache__` exists and should not be treated as source.
- `profileInterpreter.Profile` loads profiles relative to the module directory, so tests should pass profile paths like `profiles/modsprofile.yaml`.
- Because there is no refactor in scope, route tests should use mocks where direct filesystem/cache behavior would make tests brittle.
- Avoid broad snapshot tests of full XML output. They will be noisy and brittle. Prefer targeted assertions on important elements and attributes.


## Suggested First Pull Request Scope

Keep the first test change small:

1. Replace `tests/test.py` with real tests for `fileSupport.py` and `profileInterpreter.py`.
2. Add one or two `Profile.convertRowToXmlString()` tests using `profiles/modsprofile.yaml`.
3. Add basic Flask GET route tests.
4. Run `uv run ./run_tests.py`.

This gives useful coverage quickly while leaving deeper upload, ZIP, and EAD tests for follow-up work.
