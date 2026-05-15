# MODS EAD maker flask

Flask app for generating MODS and EAD XML files from spreadsheet metadata.

## Contents

- [Overview](#overview)
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

The main Flask entry point is `flask_app.py`. Core spreadsheet, XML, ZIP, preview, and filename helpers live in `fileSupport.py`. YAML profile interpretation and MODS XML generation live in `profileInterpreter.py`. Legacy EAD/MODS code lives under `legacy/`.


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


## Spreadsheet demos

Demo spreadsheets are available in `spreadsheet_demos/`.

To try one:

1. Start the Flask development server.
2. Open the route listed in `spreadsheet_demos/README.md`.
3. Upload the matching `.xlsx` file.
4. Select the listed sheet.
5. Preview the generated MODS XML or download the ZIP of `.mods.xml` files.

The demos include basic MODS records, repeated names/subjects, TIFF-oriented image records, John Nicholas Brown Center syllabi records, and music thesis records.


## Running tests

Run the full unittest suite:

```sh
uv run ./run_tests.py
```

Run a specific test module, class, or method:

```sh
uv run ./run_tests.py tests.test
uv run ./run_tests.py tests.test.TestMain
uv run ./run_tests.py tests.test.TestMain.test_name
```

Increase test output verbosity:

```sh
uv run ./run_tests.py --verbose
```

The test runner exists, but the test suite may still contain incomplete or template tests while `uv` support and project modernization are underway.
