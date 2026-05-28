# MODS EAD maker flask

Flask app for generating MODS and EAD XML files from spreadsheet metadata.

## Contents

- [Overview](#overview)
- [How the MODS Maker works: brief overview](#how-the-mods-maker-works-brief-overview)
- [How the MODS Maker works: more info](#how-the-mods-maker-works-more-info)
- [Requirements](#requirements)
- [Local installation](#local-installation)
- [Local usage](#local-usage)
- [Spreadsheet demos](#spreadsheet-demos)
- [Running tests](#running-tests)


## Overview

The app provides browser-based tools for:

- Uploading `.xlsx` spreadsheets and generating MODS XML files using YAML profiles in `profiles/`.
- Previewing generated MODS output before download.
- Packaging generated MODS files into ZIP downloads.
- Uploading `.xlsx` spreadsheets and generating EAD XML through the legacy EAD maker.
- Viewing and downloading YAML metadata profiles.
- Filling out profile-based forms that generate MODS XML.

The main Flask entry point is `flask_app.py`. Core spreadsheet, XML, ZIP, preview, and filename helpers live in `lib/fileSupport.py`. YAML profile interpretation and MODS XML generation live in `lib/profileInterpreter.py`. Legacy EAD/MODS code lives under `legacy/`.


## How the MODS Maker works: brief overview

The MODS Maker turns spreadsheet rows into MODS XML files. 

First, users choose a url that determines which mapping-profile to apply to the spreadsheet. For example, `/modsmaker/modsprofile` will apply the `modsprofile.yaml` profile to the uploaded spreadsheet.

Then, users upload an `.xlsx` file, select a sheet, preview the generated XML, and download a ZIP of `.mods.xml` files.

The important idea is that the spreadsheet does not directly define the XML structure. The YAML profile does. Each profile in `profiles/` describes which spreadsheet columns to read and how those values should become MODS elements, attributes, filenames, repeated fields, names, subjects, rights statements, and other metadata.

For image records, include a `noteImageAltText` column. When `typeOfResource` is `still image`, the MODS Maker requires that column to stay within the configured character limit and writes it as `<mods:note type="image_accessibility_alt_text">...</mods:note>`.

The default alt-text limit comes from the active YAML profile, currently 250 characters. To override it without changing code or profiles, add this setting to the `.env` file in the parent directory of this repository:

```sh
IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS=250
```

By default, MODS validation errors stop Preview and Download. To inspect generated MODS despite validation failures, uncheck `Enforce validations`. The app will continue processing and show validation warnings, but generated MODS may not work in the Workshop.

_(EAD documentation to come)_


## How the MODS Maker works: more info

The MODS Maker is a spreadsheet-to-MODS-XML generator where the YAML profile defines the mapping rules.

At a high level:

1. User opens a profile-specific route, for example:

```text
/modsmaker/modsprofile
/modsmaker/musictheses
/modsmaker/jnbcsyllabi
```

2. `flask_app.py` loads the matching YAML file:

```text
profiles/modsprofile.yaml
profiles/musictheses.yaml
profiles/jnbcsyllabi.yaml
```

The route name maps directly to the YAML filename.

3. User uploads an `.xlsx` spreadsheet and chooses a sheet.

4. `lib/fileSupport.py` reads the spreadsheet rows into dictionaries:

```python
{
    'identifierFileName': 'demo_image_0001',
    'fileTitle': 'Front entrance',
    'typeOfResource': 'still image',
}
```

5. `lib.profileInterpreter.Profile` loads the YAML profile and uses it to decide:

- what the MODS root element should be
- which spreadsheet column names to read
- which MODS elements to create
- which attributes to add
- how repeated fields are split
- how names, roles, URIs, dates, and subjects are parsed
- which rows to skip
- which filename column to use
- what file extension to append

For example, in `modsprofile.yaml`:

```yaml
filenamecolumn: identifierFileName
fileextension: ".mods.xml"
```

means a row with:

```text
identifierFileName = demo_image_0001
```

generates:

```text
demo_image_0001.mods.xml
```

The `fields:` section is the core mapping. A simplified example:

```yaml
fields:
  - type: element
    name: titleInfo
    children:
      - type: element
        name: title
        text:
          - type: value
            values:
              - {type: col, header: fileTitle, method: value}
```

That says: create a MODS `<titleInfo>` element, then a child `<title>`, and fill its text from the spreadsheet column `fileTitle`.

So this row:

```text
fileTitle = Front entrance of the sample building
```

becomes roughly:

```xml
<mods:titleInfo>
  <mods:title>Front entrance of the sample building</mods:title>
</mods:titleInfo>
```

Repeating fields are also profile-driven. A YAML block can say "read this column, split multiple values, parse each entry, and create one MODS element per entry." That is how creator names, subjects, genres, and authority URIs get expanded.

The Flask app itself does not contain much MODS logic. It mostly handles upload, preview, and download. The YAML profile is where the metadata model lives, and `lib/profileInterpreter.py` is the engine that interprets that profile into XML.

_(EAD documentation to come)_


## Requirements

- [uv](https://docs.astral.sh/uv/#installation)


## Local installation

```sh
cd /path/to/mods-ead-maker-flask-stuff/
git clone git@github.com:Brown-University-Library/mods-ead-maker-flask.git
cd mods-ead-maker-flask
uv sync --locked
```

To run commands inside the managed environment:

```sh
uv run flask --version
```


## Local usage

Start the Flask development server:

```sh
FLASK_APP=flask_app.py FLASK_ENV=development FLASK_DEBUG=1 uv run flask run
```

This starts `flask_app.py` in debug mode inside the `uv` environment. The environment-variable form is used because the project pins Flask 2.0.x.

By default, Flask serves the app at:

```text
http://127.0.0.1:5000
```

Useful local routes include:

- `http://127.0.0.1:5000/modsmaker`
- `http://127.0.0.1:5000/modsmaker/<profile-name>`
- `http://127.0.0.1:5000/eadmaker`
- `http://127.0.0.1:5000/profiles/`
- `http://127.0.0.1:5000/forms/`
- `http://127.0.0.1:5000/resources`

For example, the default MODS Maker route redirects to the `modsprofile` profile:

```text
http://127.0.0.1:5000/modsmaker/modsprofile
```

For MODS spreadsheet uploads, the route URL selects the mapping profile. For example, `/modsmaker/musictheses` uses `profiles/musictheses.yaml`. The upload form does not choose or detect a profile from the spreadsheet file.

The upload form checks `Enforce validations` by default. Leave it checked to stop Preview and Download when validation errors are found. Uncheck it to continue processing while showing validation warnings.


## Spreadsheet demos

Demo spreadsheets are available in `spreadsheet_demos/`.

To try one:

1. Start the Flask development server.
2. Open the route listed in `spreadsheet_demos/README.md`.
3. Upload the matching `.xlsx` file.
4. Select the listed sheet.
5. Preview the generated MODS XML or download the ZIP of `.mods.xml` files.

The listed route is part of the demo setup: it determines which YAML mapping profile will be applied to the uploaded spreadsheet.

The demos include basic MODS records, repeated names/subjects, TIFF-oriented image records, John Nicholas Brown Center syllabi records, and music thesis records.


## Running tests

Run the full unittest suite:

```sh
uv run ./run_tests.py
```

Run a specific test module, class, or method:

```sh
uv run ./run_tests.py tests.test_file_support
uv run ./run_tests.py tests.test_file_support.TestFileSupport
uv run ./run_tests.py tests.test_file_support.TestFileSupport.test_clean_string_for_filename_removes_invalid_characters
```

Increase test output verbosity:

```sh
uv run ./run_tests.py --verbose
```

---
