# Doc Scanner Mover

[Download the updated script ZIP](https://github.com/seanpconners-source/Doc-Scanner-Mover/raw/refs/heads/main/downloads/Doc-Scanner-Mover-Updated.zip).
It includes the Python script, dependency list, and replacement instructions.
Use the updated Excel export you already have; spreadsheets and scans are not
included in the repository download.

`Check and Rename Incoming Scans.py` watches completed checklist scans and names
them `YYYYMMDD-ITEM-DESCRIPTION-PO005157_1.pdf` (or `STOCK`). Existing names get a
new occurrence number; existing scans are never overwritten.

The original script's "File Rename Error" notices actually meant that its OCR
did not find an item number or PO. Testing the four supplied scans reproduced
missing fields, a false STOCK result, and a wrong item number. Tesseract's default
page segmentation skips text inside this form's table, broad crops misread the
header, and a pen stroke interferes with one printed item number. The old STOCK
test also accepts body instructions containing "stock" when it misses the PO.
Its image filtering and extra PDF rasterization reduce clarity further.

## Windows installation

Use Python 3.10 or newer. In PowerShell, from this repository:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe '.\Check and Rename Incoming Scans.py' --check
```

To keep an **existing shortcut**, stop the old monitor, back up its script, and
copy the replacement to exactly the old script's filename and location. Install
`requirements.txt` into the **same Python installation used by the shortcut**;
creating a separate virtual environment does not install packages into that
installation. Check the shortcut's Properties → Target to identify its Python
and script paths. If it uses `pythonw.exe`, use `python.exe` from the same Python
installation for installation and `--check`. Then start the shortcut once.

Keep the installed Tesseract OCR and Poppler from the original setup. The script
uses the original Windows executable locations if they exist, otherwise PATH.
If installed elsewhere, set `SCAN_TESSERACT_PATH` to `tesseract.exe` and
`SCAN_POPPLER_PATH` to the directory containing `pdfinfo.exe` and `pdftoppm.exe`.
Tesseract's English language data must be installed. `--check` reports missing
tools, inaccessible folders, invalid Excel data, and output write failures
without processing scans.

The original input/output/debug/workbook paths remain the Windows defaults.
Override them with `--input`, `--output`, `--debug`, and `--items`, or environment
variables `SCAN_INPUT_FOLDER`, `SCAN_OUTPUT_FOLDER`, `SCAN_DEBUG_FOLDER`, and
`SCAN_ITEM_LIST`. The workbook's **first worksheet** must have item numbers in
column A and descriptions in column B. Headers are allowed. Numeric cells,
leading zeroes, and text such as `13200091.0` are normalized. Missing, empty, or
conflicting item lists stop processing instead of silently returning UNKNOWN.
The monitor reloads the list when the workbook changes.

## Manually refresh the item list

The item export must still be pulled manually from your source system. After
each pull, back up the previous file and save the refreshed export as:

```text
C:\Users\sconners\Documents\Flavor List\FlavorList.xlsx
```

Keep `No.` / item number in column A and `Description` in column B of the first
worksheet. Both supplied exports already have this supported format. A new
file named `FlavorList - UPDATED 2026-10-07.xlsx` placed alongside the old file
does not change which workbook is loaded. Save it under the configured filename
above, or explicitly select its actual path with `--items` / `SCAN_ITEM_LIST`.

The revised monitor reloads a replaced workbook automatically. The original
uploaded script loads it only once at startup, so that version needs a restart
after replacement. Use the revised script and `--check` to verify the selected
workbook before processing production scans.

Comparison of the two supplied exports found 2,663 versus 2,946 eight-digit
item numbers: 283 added, none removed, and 1,232 changed descriptions (many
punctuation changes). The old list lacks `13100746`, `13100747`, `13100748`, and
`13100755`; the updated list includes them. Stale catalog entries explain those
UNKNOWN names. The espresso `13100029` and chocolate `13100228` entries already
exist in both lists; their sample failures came from OCR. Workbook freshness
does not by itself fix missing POs or incorrectly read item numbers.

Mapped drives such as `P:` must be available to the Windows account running the
script. A scheduled task running as a different account may need a UNC share
path instead. Supply your actual share path using the supported overrides.

## Run

**Stop the old monitor before starting this version.** First test with copies
of several known scans in separate input/output folders and a copy of the real
workbook. Run `--check` using those paths, then `--once`:

```powershell
.\.venv\Scripts\python.exe '.\Check and Rename Incoming Scans.py' --once --input 'C:\ScanTest\Incoming' --output 'C:\ScanTest\Sorted' --debug 'C:\ScanTest\Debug' --items 'C:\ScanTest\FlavorList.xlsx'
```

After checking the names against the documents, start continuous monitoring:

```powershell
.\.venv\Scripts\python.exe '.\Check and Rename Incoming Scans.py'
```

The monitor waits for a file's size and modification time to stay unchanged for
10 seconds before processing. Failed reads/conversions leave the source in the
input folder and retry after 60 seconds. This avoids reading scans while the
scanner is still copying them. A very long pause during a scanner upload can
look like completion; increase `--settle-seconds` for that scanner if necessary.

PDFs retain their original bytes and **all pages**; only the first page is used
to identify the checklist. PNG/JPEG inputs are actually converted to PDF, with
image handles closed before removing the source on Windows. Sorting copies
the complete output successfully before removing the source.

## Recognition and diagnostics

Existing PDF text is used when available. Otherwise the original first page is
rendered once at 300 DPI. OCR first reads the whole page. When it detects the
Production Checklist form, it reads padded value cells inside the header table,
excluding the bordering rules. Broad crops, contrast adjustment, alternate page
segmentation, and rotations remain fallbacks for tilted or shifted layouts.
A destructive fixed black/white threshold is avoided.

Item number corrections such as `O` → `0` and `I`/`J` → `1` require a match in the
Excel list. Clearly labeled exact item numbers absent from the list retain
their number and use `UNKNOWN`, with a separate catalog warning. Ambiguous
numbers and missing fields use `NO_CODE_FOUND` / `NOPO` for manual review;
the program does not guess among conflicting readings. Six-digit PO forms
such as `PO005157`, `PO: 005157`, `P.O. # 005157`, and spaced digits are supported.
PO readings permit common character confusions such as `S` → `5` only in the
six-character numeric field with at least four actual digits. Focused PO cell
readings take precedence over body text. On Production Checklists, `STOCK` must
be in the PO cell or explicitly labeled as the PO, rather than in body text.

Each attempt gets a separate directory in the debug folder with `ocr.txt`, the
raw text of each attempted pass, the selected fields, the output, and the
specific review/error reason. Image OCR also saves a first-page PNG and any
attempted field crops. Use these files to distinguish unreadable OCR from a
missing workbook item. They contain document data; retain/delete them according
to your document retention policy.

`--once` exit codes: **0** = all scanned files processed without review (or no
eligible files); **1** = setup/processing failure or a file still changing;
**2** = files processed but at least one needs OCR/catalog review.
`--check` validates prerequisites, not recognition of a document.

## Email

The password embedded in the uploaded script has been removed. Rotate that
password and provide the replacement through `SCAN_SMTP_PASSWORD` in the
environment of the Windows process. Do not put it in this script or Git.
If this variable is absent, email is disabled and warnings remain in the logs.

The original Zoho host, sender, and recipient remain the defaults. Optional
overrides: `SCAN_SMTP_HOST`, `SCAN_SMTP_PORT`, `SCAN_SMTP_USERNAME`,
`SCAN_EMAIL_FROM`, and `SCAN_EMAIL_TO`. STARTTLS certificate verification remains
enabled. A completed scan that needs review produces one notice with the actual
reason and diagnostics path. An SMTP failure is logged and never retries a
completed move. No emails are sent during tests.

## Development and validation

Use the existing checkout; cloud tasks are already isolated and do not need a
Git worktree. Install Tesseract and Poppler through your OS package manager if
not already present. On the prepared cloud instance both are installed.

```bash
python -m venv /workspace/.venvs/doc-scanner-mover
/workspace/.venvs/doc-scanner-mover/bin/python -m pip install -r requirements-dev.txt
/workspace/.venvs/doc-scanner-mover/bin/python -m pytest -q
```

The regression suite exercises parsing, Excel normalization, ambiguous fields,
duplicate names, full PDF preservation, true image-to-PDF conversion, changing
uploads, processing failures, and real Tesseract recognition of generated scans
with shifted fields, faint text, and rotations. These tests validate the cloud
implementation; Windows network-drive access, the production Excel list, and
live SMTP delivery require validation on the owner's Windows machine.

Private document regressions are available separately:

```bash
SCAN_TEST_SAMPLE_DIR=/path/to/private/scans /workspace/.venvs/doc-scanner-mover/bin/python -m pytest -q
```

Use the original filenames of the four supplied PDFs in that directory. The
tests cover each PDF and an upside-down copy of one checklist. Without the
variable, these five tests are skipped. Uploaded PDFs and OCR diagnostics
remain outside the repository. Add `SCAN_TEST_ITEM_LIST=/path/to/FlavorList.xlsx`
to use the actual workbook; otherwise the tests use a sample catalog transcribed
from printed document headers. On the prepared cloud instance the supplied
updated workbook is `/workspace/library-files/doc-scanner-mover/FlavorList.xlsx`;
the old export is preserved alongside it as `FlavorList-original.xlsx`.

The verified readings are:

| Uploaded filename identifier | Item number | PO | Pages preserved |
| --- | --- | --- | --- |
| `13100747-UNKNOWN-NOPO` | `13100747` | `PO005157` | 8 |
| `NO_CODE_FOUND-UNKNOWN-STOCK` | `13100029` | `PO005191` | 8 |
| `NO_CODE_FOUND-UNKNOWN-PO005157` | `13100748` | `PO005157` | 7 |
| `43100228-UNKNOWN-NOPO` | `13100228` | `PO031440` | 8 |

Each real-document test verifies exact output bytes, page count, original sample
preservation, and no OCR review/email request. All four supplied scans were also
validated using the supplied updated production workbook, including complete
descriptions. They pass without missing fields, UNKNOWN names, or review notices.
The CLI preflight, actual batch, and repeated empty batch were verified with
that workbook. Windows deployment and live SMTP delivery remain untested here.
