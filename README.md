# Personalized CSV Creator

A desktop GUI utility that converts Excel workbooks (`.xlsx`/`.xls`) into
comma-separated value (CSV) files while giving you fine-grained control over
formatting, encoding, quoting, and line endings. The application uses Tkinter
for the interface and pandas/openpyxl for Excel parsing, and persists your last
used settings for quick repeat conversions.

## Features

- **Excel input** – Browse to any `.xlsx` or `.xls` file and pick the sheet to
  convert.
- **Custom separators** – Choose from common delimiters (comma, semicolon, tab,
  pipe, colon) or supply any custom string, including multi-character or
  Unicode separators (for example `||`, `---`, or `😊`).
- **Flexible quoting** – Three quoting modes (none, text fields only, all
  fields) paired with selectable quote characters (`"`, `'`, `` ` ``, or any
  custom single character). Values that contain the separator, a newline, or the
  quote character are always quoted to keep the CSV valid.
- **Escaping strategy** – When a value contains the quote character, the app
  doubles that character (e.g. `"He said ""hi"""`) so the output remains
  standards-compliant.
- **Preview** – Displays the first 20 rows rendered with the currently selected
  options so you can verify separators and quoting before saving.
- **Encoding & newline control** – Export using UTF-8 (with or without BOM),
  ISO-8859-1, or Windows-1252, and choose OS-default, Unix (`\n`), or Windows
  (`\r\n`) line endings.
- **Threaded conversion with progress** – Large sheets are converted on a worker
  thread. The status bar reports progress in 1,000-row increments so the UI
  remains responsive even for 50k+ rows.
- **Robust error handling** – Clear dialogs are shown for unreadable files,
  protected sheets, invalid separators/quotes, or write/permission issues.
- **Configuration persistence** – Last-used settings (including custom
  separators, quote character, encoding, and line-ending choices) are saved to
  `~/.personalized_csv_creator_config.json` for convenience.

## Requirements

- Python 3.7.2 or later within the 3.7 series
- `pandas` (tested with 1.3.5)
- `openpyxl` (for `.xlsx` support)
- `xlrd` 1.2.0 (for legacy `.xls` files)

Install the dependencies with:

```bash
pip install -r requirements.txt
```

## Running the app

```bash
python app.py
```

1. Click **Browse…** to select an Excel workbook.
2. Choose the worksheet, separator, quoting mode/character, encoding, and line
   ending.
3. Review the live preview and adjust settings if needed.
4. Press **Save CSV…** to pick an output path. Conversion runs in the background
   with progress updates.

### Quoting behaviour in detail

- **No quoting:** Text fields are emitted as-is, but values containing the
  separator, newline characters, or the quote character are still quoted to keep
  the CSV valid.
- **Quote text fields:** Text (non-numeric, non-date) values are wrapped with the
  chosen quote character. Numbers, dates (`YYYY-MM-DD`), and booleans remain
  unquoted unless quoting is required to escape special characters.
- **Quote all fields:** Every value is quoted.
- In all modes, the quote character inside a field is escaped by doubling it.

Dates detected via pandas are exported as `YYYY-MM-DD` when they contain no time
component. Datetimes keep their ISO 8601 representation (including timezone
information when present). Empty cells become empty strings.

### Configuration storage

Preferences are saved automatically to
`~/.personalized_csv_creator_config.json` when you close the application. Delete
that file to reset the app to default settings.

## Packaging with PyInstaller (optional)

Use the included helper script to produce a standalone executable:

1. Install PyInstaller (preferably in a virtual environment):
   ```bash
   pip install pyinstaller
   ```
2. Run the build helper (this defaults to a single-file, windowed executable):
   ```bash
   python build_executable.py
   ```
   Use `--console` to keep the terminal window, `--onedir` to create a folder
   bundle, or `--icon path/to/icon.ico` to embed a custom icon.
3. Collect the packaged binary from the `dist/` directory. The default output is
   `dist/PersonalizedCSVCreator` (or `PersonalizedCSVCreator.exe` on Windows).

The script automatically cleans previous PyInstaller artifacts so repeated
builds stay reliable. If you prefer to call PyInstaller manually, the script
emits the equivalent of:

```bash
pyinstaller --name PersonalizedCSVCreator --onefile --noconsole app.py
```

PyInstaller bundles pandas and openpyxl automatically. Adjust the helper
arguments to add icons or switch packaging modes as needed.

## Troubleshooting

- Ensure the Excel file is not password-protected and is accessible with your
  current permissions.
- Multi-character separators and Unicode quote characters are supported, but
  custom quote characters must be a single Unicode codepoint.
- UTF-8 with BOM is recommended when targeting legacy Excel installations that
  expect BOM-prefixed CSV files.

Enjoy creating perfectly formatted CSVs tailored to your workflow!
