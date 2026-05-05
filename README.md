# ReadBubbles

ReadBubbles converts bubble-drawing PDFs into standardized Excel workbooks for position tracking, audit preparation, and PPAP follow-up.

The project is built for a workflow where a drawing contains numbered inspection bubbles, and those position numbers need to be collected into a repeatable Excel format. Instead of rebuilding that workbook by hand for every revision, the tool extracts positions from the PDF, reuses information from earlier revisions when possible, and generates a formatted workbook with helper sheets.

## Overview

ReadBubbles takes a bubble drawing PDF as input and creates an Excel workbook in the same folder. The generated workbook includes:

- A revision-specific worksheet containing the extracted position numbers
- A cover sheet
- A PPAP worksheet
- A change notes worksheet

The project is aimed at quality or engineering workflows where consistency between drawing revisions matters and where manual Excel preparation is time-consuming.

## Features

- Extracts part number and revision from PDF text
- Detects and cleans position numbers from drawing text
- Expands ranges such as `0103-0106` into individual positions
- Categorizes position numbers by type using configurable rules
- Loads data from a previous workbook revision when available
- Preserves previously entered manual information where possible
- Adds workbook formatting, validation lists, formulas, and summary counts
- Creates support sheets for cover information, PPAP work, and change tracking
- Supports packaging as a Windows executable

## How It Works

### High-Level Flow

1. The user provides the path to a bubble drawing PDF.
2. Text is extracted from the PDF using `pdfminer`.
3. Candidate position numbers are cleaned and normalized.
4. Each position number is categorized based on rules in `settings.json`.
5. If an older workbook exists, the latest prior revision sheet is loaded.
6. A new revision sheet is written to Excel.
7. Formatting, data validation, counts, and helper sheets are added.

### Main Modules

- [main.py]: Entry point that orchestrates the full workflow.
- [file_handler.py]: Handles PDF path input and validation.
- [pdf_extract.py]: Extracts text from PDFs and cleans position-number data.
- [pos_number.py]: Converts a position number into a row template with a derived type.
- [pos_numbers.py]: Builds a sorted collection of position-number rows.
- [dataframe_handler.py]: Creates and merges the main pandas dataframe.
- [df_add_gears.py]: Adds special gear-related parameter rows.
- [xl_handler.py]: Creates or loads the target workbook and reads older revision data.
- [xl_formater.py]: Applies worksheet styling, validations, formulas, and helper-sheet generation.
- [xl_cover.py]: Builds the cover sheet.
- [xl_ppap.py]: Builds the PPAP sheet.
- [xl_change_notes.py]: Builds the change notes sheet.
- [settings.json]: Central configuration for parsing, columns, validation lists, colors, and sheet behavior.

## Requirements

This project can be run either from source or as a packaged executable.

Recommended environment:

- Python 3.10 or newer
- Microsoft Excel-compatible `.xlsx` workflow
- Windows for the packaged `.exe` flow

Python packages used by the project:

- `pandas`
- `openpyxl`
- `pdfminer.six`
- `colorama`
- `cx_Freeze`

If you plan to run from source, install the dependencies into a virtual environment.

## Installation

### Run From Source

Create and activate a virtual environment, then install the required packages:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl pdfminer.six colorama
```

### Build Executable

To build a Windows executable with `cx_Freeze`:

```bash
python setup.py build
```

The build process includes:

- `settings.json`
- `scania-symbol.png`

## Usage

### Basic Usage

You can run the program interactively:

```bash
python main.py
```

The tool will prompt for the path to a PDF:

```text
Please input path to pdf:
```

You can also pass the PDF path directly:

```bash
python main.py "C:\path\to\drawing.pdf"
```

### What the Program Does

When the program runs, it will:

- Validate the input path
- Extract drawing text from the PDF
- Find and clean position numbers
- Detect the part number and revision
- Create or update the output workbook
- Attempt to reuse data from an earlier revision
- Print lists of new, removed, and duplicate positions
- Save the formatted workbook

### Output

The output workbook is saved in the same folder as the input PDF.

Workbook naming format:

```text
<part number> Positionsdatablad.xlsx
```

Revision sheet naming format:

```text
<part number>_<revision>
```

Example:

- Input PDF: `1234567 Rev_3 drawing.pdf`
- Output workbook: `1234567 Positionsdatablad.xlsx`
- Revision sheet: `1234567_3`

## Configuration

### `settings.json`

The project is heavily configuration-driven. Many of the business rules live in [settings.json](C:/Users/jxqac9/.codex/worktrees/d23e/ReadBubbles/settings.json:1) rather than in hardcoded Python logic.

Important sections include:

- `PosTypes`: Regex patterns used to classify position numbers
- `Columns`: Default columns and default values for the main dataframe
- `Column Size`: Column widths used in generated worksheets
- `Data Validation`: Excel dropdown values for the main worksheet
- `Conditional`: Conditional formatting rules for the main worksheet
- `Change Notes Validation`: Dropdown values for the change notes sheet
- `Colors`: Shared color definitions
- `PPAP`: PPAP sheet columns, validation, chart, and conditional settings
- `Cover Settings`: Labels and layout for the cover sheet
- `Extract Settings`: `pdfminer` layout tuning values
- `Zoom`: Worksheet zoom defaults

### When to Change Settings

You may need to update `settings.json` when:

- Adding a new position-number type
- Changing the set of workbook columns
- Updating dropdown values used by the team
- Adjusting conditional formatting colors
- Tuning PDF extraction for a different drawing layout
- Changing PPAP or cover-sheet formatting

Because configuration affects parsing and workbook output, changes to `settings.json` should be treated like code changes and ideally be regression-tested.

## Excel Workbook Structure

### Main Revision Sheet

This is the primary worksheet generated for the current part number and revision. It contains one row per position number, along with metadata columns such as:

- `Position Number`
- `Type`
- `Specification`
- `Gear ID`
- `Classification`
- `Audited`
- `Measurement Instrument`
- `Comment`
- `CMS Audit`
- `MSA1`
- `MSA2/3`
- `Manually Added`
- `Nominal`
- `UTL`
- `LTL`
- `Exemption Approved Date`
- `Exemption Valid To`

The sheet also includes:

- Excel table formatting
- Conditional formatting
- Data validation dropdowns
- Audit count formulas
- Classification counts
- Freeze panes
- Column sizing and borders

### Cover Sheet

The `Cover` sheet acts as a front page for the workbook. It includes:

- Part number
- Revision
- Informational text
- Signature areas
- Embedded logo

The part number and revision are updated automatically when the workbook is generated.

### PPAP Sheet

The PPAP sheet provides a secondary view intended for PPAP-related tracking and action management. It includes:

- Position numbers copied from the main revision sheet
- Selected values copied by formula from the main worksheet
- Validation lists for PPAP fields
- Conditional formatting
- A simple chart for `Ppk/Cpk`
- Action counts

The sheet is created automatically and is hidden by default after generation.

### Change Notes Sheet

The `Change_Notes` sheet provides a standard place to document revision changes. It includes columns such as:

- `Date`
- `Rev`
- `Position Number`
- `Change Type`
- `Description`
- `SSSID`
- `Verified by ID`
- `Verified Date`
- `Comment`

This sheet is intended to support traceability across revisions.

## Position Parsing Rules

The project includes logic to normalize several drawing notations into usable position-number rows. Examples include:

- Single positions such as `0101`
- Ranges such as `0103-0106`
- Slash-separated lists such as `0101/0102`
- Decimal variants such as `1101.1/1102.1`
- Gear-related markers such as `0200-0299`

The extraction logic also excludes some values that look numeric but should not become inspection positions, such as:

- Dates
- Some long numeric identifiers
- Certain thread or special-code patterns

This logic lives mostly in [pdf_extract.py](C:/Users/jxqac9/.codex/worktrees/d23e/ReadBubbles/pdf_extract.py:97).

## Revision Reuse Behavior

If a workbook already exists for the same part number, the tool searches for the latest earlier revision sheet and loads it. It then:

- Aligns old and new columns if the schema changed
- Reuses old row values for positions that still exist
- Reports newly added positions
- Reports removed positions
- Re-adds rows marked as manually added

This logic lives primarily in [xl_handler.py](C:/Users/jxqac9/.codex/worktrees/d23e/ReadBubbles/xl_handler.py:49) and [dataframe_handler.py](C:/Users/jxqac9/.codex/worktrees/d23e/ReadBubbles/dataframe_handler.py:80).

## Limitations

- Output quality depends on how well `pdfminer` can extract text from the source PDF.
- Rotated or unusually arranged drawing text may not be fully supported.
- Some gear-related workflows require manual user input.
- Parsing rules appear tailored to a specific engineering or quality domain.
- There is currently no bundled test suite in the repository.
- Error handling is fairly minimal in some places, especially around Excel and revision loading.

## Testing

The repository currently does not include automated tests.

## Troubleshooting

### The PDF Path Is Rejected

Check that:

- The file exists
- The path points to a file, not a folder
- The file extension is `.pdf`

### Part Number or Revision Is Missing

The extraction logic expects certain text patterns in the PDF. If the drawing format differs from the expected layout, part and revision detection may fail.

Review:

- The drawing text structure
- The revision label format
- The extraction tuning values in `Extract Settings`

### Position Numbers Are Missing or Incorrect

This is usually caused by PDF text extraction layout issues or a drawing format that does not match the current parsing rules.

Things to check:

- Whether the PDF contains selectable text
- Whether the values appear in an unusual orientation
- Whether the parsing rules in `pdf_extract.py` need to be extended
- Whether `Extract Settings` in `settings.json` need adjustment

### Workbook Data Did Not Carry Forward From an Earlier Revision

Check that:

- The output workbook already exists
- The previous sheet naming matches `<part number>_<revision>`
- The older revision sheet contains the expected columns

### Import or Dependency Errors

Make sure all required packages are installed in the active Python environment:

```bash
pip install pandas openpyxl pdfminer.six colorama cx_Freeze
```

## Contributing

If you change parsing logic, workbook structure, or configuration:

- Add or update tests for the affected behavior
- Keep `settings.json` and the README in sync
- Prefer adding regression cases for tricky drawing formats
- Be careful when changing workbook column names, since revision reuse depends on them

For larger changes, it is a good idea to test against a small set of known PDFs and compare the generated workbook against expected output.

## Future Improvements

- Add an automated test suite
- Add sample fixtures for parsing and workbook generation
- Add clearer error messages and structured logging
- Add sample screenshots or example output documentation

## License

Intended for internal use only.

## Ownership

Contact JXQAC9 for workflow questions, rule updates, or bug fixes.
