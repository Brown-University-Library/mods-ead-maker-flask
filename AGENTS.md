# AGENTS.md - Repository Agent Instructions

This file defines the coding guidance for LLM coding agents working in this repository.
When these instructions conflict with older IDE, Copilot, or contributor notes, prefer this file.


## Project Basics

- Primary language: Python.
- Application framework: Flask.
- App entry point: `flask_app.py`.
- Production-compatible Python runtime: Python `>=3.8,<3.9`, matching `pyproject.toml`.
- Dependency pins in `pyproject.toml` reflect the versions currently expected for production compatibility.
- Project root is the directory containing this file, `.git/`, `pyproject.toml`, and `flask_app.py`.


## Current Tooling Status

- This project is being prepared to use `uv`; prefer `uv` commands when the environment supports them.
- `run_tests.py` exists and is intended to be the test runner, but the test suite may still contain template or incomplete tests.
- If `uv` is unavailable in the current environment, state that clearly before using a local fallback.
- Do not assume there is a `main.py`; this is a Flask app centered on `flask_app.py`.


## How To Run

- Assume commands run from the project root.
- Start the development Flask app with:
  - `./runflaskappindebug.sh`
- Run all tests with:
  - `uv run ./run_tests.py`
- Run a narrower unittest target with:
  - `uv run ./run_tests.py tests.test`
  - `uv run ./run_tests.py tests.test.TestMain`
  - `uv run ./run_tests.py tests.test.TestMain.test_name`


## Coding Directives

### Python Compatibility

- Write code compatible with Python 3.8.
- Do not use Python 3.9+ only syntax such as builtin generic type annotations (`list[str]`, `dict[str, int]`).
- Do not use Python 3.10+ only syntax such as PEP 604 unions (`str | None`) or `match` statements.
- If adding type hints, use Python 3.8-compatible forms such as `typing.List`, `typing.Dict`, and `typing.Optional`.
- Keep dependency choices compatible with the pinned package versions in `pyproject.toml`.


### Style

- Inspect `ruff.toml` before broad edits.
- Current formatting expectations include:
  - max line length: 125
  - indentation: 4 spaces
  - quote style: single quotes
  - Ruff target version: `py38`
- Favor clarity and explicitness over cleverness.
- Prefer small, focused functions with straightforward control flow.
- Do not define functions inside other functions unless there is a specific local reason.


### Flask Architecture

- `flask_app.py` contains the Flask route handlers.
- Route handlers should act as managers:
  - Parse request input.
  - Perform minimal validation and shaping.
  - Delegate MODS, EAD, profile, Excel, XML, and file-generation work to helper modules.
  - Convert returned values into Flask responses, redirects, rendered templates, downloads, or JSON.
- Keep reusable domain logic out of route handlers where practical.
- Prefer pure helper functions that accept plain Python values instead of Flask request objects.
- Existing helper modules include:
  - `fileSupport.py` for spreadsheet, XML, ZIP, preview, and file-output helpers.
  - `profileInterpreter.py` for YAML profile interpretation and MODS XML generation.
  - `legacy/EADMaker.py` and `legacy/MODSMaker.py` for legacy EAD/MODS behavior.
- Do not add Django conventions, Django management commands, or Django directory assumptions to this project.


### Templates And Profiles

- HTML templates live under `templates/`.
- YAML metadata profiles live under `profiles/`.
- Preserve existing profile behavior unless the task explicitly changes profile semantics.
- Be careful with XML output formatting and filenames; these are user-facing export behavior.


### Tests

- Use the standard library `unittest` framework.
- Add or update focused tests for behavior changes when practical.
- Prefer tests that cover:
  - the happy path
  - at least one failure or edge case
- Run `uv run ./run_tests.py` after changes when the environment supports it.
- If tests cannot be run, state the exact command that should be run and why it was not run.


## Change Workflow Expectations

When implementing a change:

1. Read relevant surrounding code and match existing conventions.
2. Make the smallest correct change that satisfies the request.
3. Keep production compatibility with Python `>=3.8,<3.9` and pinned dependencies.
4. Update tests when the change affects behavior.
5. Run `uv run ./run_tests.py` when feasible.


## If Instructions Are Missing Or Ambiguous

- Do not ask questions unless the ambiguity blocks safe progress.
- Make reasonable assumptions, state them explicitly, and continue.
- If blocked, provide:
  - what you tried
  - what you found in the repo
  - a concrete next step
