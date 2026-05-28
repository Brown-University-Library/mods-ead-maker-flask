Prompt:

Goal: move the "250" value of the 250-character image-description-alt-text validation limit/warning into a `.env` setting.

Context:

- currently if one wanted to change the 250-character validation-limit, one would have to change the code/profile.
- i want to be able, instead, to change a setting.
- i want the code to check the `.env` setting.
- NOTE: the `.env` file does _NOT_ live in the `.git/` directory, but in its parent directory.
- if the `.env` doesn't exist, it should use the value in the profile.
- if the `.env` setting does exist, it should use the value in the `.env` setting.

Tasks:

- Review `mods-ead-maker-flask/README.md` enough to understand the overall goal of this webapp.
- Review `mods-ead-maker-flask/AGENTS.md` to understand coding-directives.
- Implement the `.env` check.

Response:

Implemented `IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS` support for the image alt-text max-character validation. The validation now reads the setting from the parent-directory `.env`, falls back to the profile `maxchars` when the file or setting is missing or invalid, and formats validation messages with the effective limit. Updated the profile messages to use `{maxchars}`, added unit coverage for fallback, override, and invalid override behavior, and documented the setting in `README.md`. Verified with `uv run ./run_tests.py`, which passed 49 tests.

---

Prompt:

Goal: Update the list of "noteImageAltText" values that will trigger an image_accessibility_alt_text validation check.

Context:

- Review `mods-ead-maker-flask/README.md` to understand the purpose of this repository, and a bit about validation.

- The check for `image_accessibility_alt_text` source data in the column "noteImageAltText" currently only occurs if the "typeOfResource" column's value is "still image".

- In addition to "still image", here is the complete list of values that should also trigger the check:
	- still image
	- books
	- images
	- journals
	- manuscripts
	- maps
	- newspapers
	- realia
	- scores
	- text_resources

Tasks:

- Review `mods-ead-maker-flask/AGENTS.md` for coding-directives.

- Update the code so that all those possible "typeOfResource" values are checked, not just the value of "still image".

- Update any tests and documentation as necessary.

Response:

Implemented the expanded `noteImageAltText` required-validation trigger list. The validation engine now supports an `in` condition, active profiles use it for `typeOfResource` values `still image`, `books`, `images`, `journals`, `manuscripts`, `maps`, `newspapers`, `realia`, `scores`, and `text_resources`, and the generated MODS note type remains `image_accessibility_alt_text`. Updated focused validation tests plus README and spreadsheet-demo documentation. Verified with `uv run ./run_tests.py`, which passed 49 tests.

---
