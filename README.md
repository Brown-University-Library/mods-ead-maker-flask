# MODS EAD Maker Flask

Flask app for generating MODS and EAD XML files from spreadsheet metadata.

The app provides browser-based tools for:

- Uploading `.xlsx` spreadsheets and generating MODS XML files using YAML profiles in `profiles/`.
- Previewing generated MODS output before download.
- Packaging generated MODS files into ZIP downloads.
- Uploading `.xlsx` spreadsheets and generating EAD XML through the legacy EAD maker.
- Viewing and downloading YAML metadata profiles.
- Filling out profile-based forms that generate MODS XML.

The main Flask entry point is `flask_app.py`. Core spreadsheet, XML, ZIP, preview, and filename helpers live in `fileSupport.py`. YAML profile interpretation and MODS XML generation live in `profileInterpreter.py`. Legacy EAD/MODS code lives under `legacy/`.


## Requirements

- Python `>=3.8,<3.9`
- `uv`

The pinned Python packages are listed in `pyproject.toml`. `uv` support is still being built out, so some local setup details may continue to change as the project is modernized.


## Local Installation With uv

This repository assumes a valid `uv.lock` file is already present. Use the lockfile as the source of truth for local installs.

From the repository root:

```sh
uv sync --locked
```

To run commands inside the managed environment:

```sh
uv run flask --version
```

Do not run `uv lock` as part of routine local setup. Only refresh the lockfile intentionally when updating project dependencies.


## Local Usage With uv

Start the Flask development server:

```sh
uv run ./runflaskappindebug.sh
```

The script sets:

- `FLASK_APP=flask_app.py`
- `FLASK_ENV=development`
- `FLASK_DEBUG=1`

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


## Running Tests With uv

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
