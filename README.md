# SAP Purchase Order Email Attacher

Windows desktop automation for processing Outlook `.msg` files, extracting purchase-order information from email and PDF content, updating an Excel delivery report, and attaching the original email to the matching SAP purchase order through SAP GUI scripting.

> **Review basis**
>
> Only the packaged `SAP_Email_Attacher.exe` was provided for this review. No standalone `.py` source files, tests, lockfile, or build configuration were available. The behavior documented here was recovered from the executable's embedded Python bytecode and package metadata using static analysis. The executable was not run during the review. Verify this README against the original source before publishing.

## Suggested project identity

- **Project name:** SAP Purchase Order Email Attacher
- **Suggested repository name:** `sap-purchase-order-email-attacher`
- **Existing application title:** SAP Email Attacher

## Project description

SAP Purchase Order Email Attacher automates a repetitive procurement operations workflow for order confirmations and other PO-related emails. A user adds Outlook `.msg` files through a Windows desktop interface or places them in the `Input` folder. The application reads the email, identifies its purchase order number, opens the matching purchase order in SAP transaction `ME23N`, and attaches the original Outlook email to that SAP purchase order. When a PDF delivery document is present, the application can also extract delivery details into an Excel report.

The application is designed to reduce the manual time previously needed to find a PO in SAP and attach an email. It also supports PDF extraction, Excel reporting, successful/failed processing queues, and dated processing logs.

The author reports that a manual attachment task previously took approximately one minute and that the automated workflow can complete one email attachment in under 10 seconds. These figures are user-reported and were not independently benchmarked during this review.

## Verified features

### Desktop interface

- Tkinter desktop window titled **SAP Email Attacher**.
- Home tab with:
  - `+ Add Emails` file picker restricted to Outlook `.msg` files.
  - `Run Now` processing button.
  - Count of emails waiting in `Input`.
  - List of files waiting to be processed.
  - Drag-and-drop support through `tkinterdnd2`.
  - Scrolling progress log showing the pipeline's console output.
- Report tab with:
  - Refreshable view of the generated Excel report.
  - Row count/status message.
  - `Open in Excel` action on Windows.
- Processing runs on a background thread so the GUI can remain responsive while the pipeline is working.

### Outlook MSG processing

- Reads `.msg` files with `extract-msg`.
- Reads subject, sender, body, and attachments.
- Tries several encodings when opening a message, including Windows-1252, CP1252, Latin-1, `chardet`, and the library default.
- Saves attachments to `Temp`.
- Uses the long attachment filename when available and falls back to the short filename.
- Sanitizes unsafe filename characters and generates a unique temporary filename when a collision exists.
- Skips an input file when a file with the same name already exists in `Processed` or `Failed`.

### Purchase-order identification

The application searches for the first value matching the verified pattern `45\d{8}` with digit boundaries. In other words, it expects a 10-digit purchase-order number beginning with `45`.

The lookup order is:

1. The email filename without its extension.
2. The email subject.
3. The email body.
4. Text extracted from PDF attachments in attachment order until a matching PO is found.

If no purchase order is found, the email is logged as failed and moved to `Failed`.

### PDF extraction and delivery report

- Opens the selected PDF with `pdfplumber`. The current attachment loop retains the last PDF encountered for report extraction.
- Extracts text from each PDF page and joins the page text.
- Builds or updates `Reports/Delivery_Report.xlsx` with the following columns:
  - Bill of Lading
  - Customer PO Number
  - Delivery Date
  - Ship To
  - Product Description
  - Customer Product
  - Quantity (Metric)
  - Customer Product Code
- Extracts report values with regular expressions for the labels and formats present in the packaged code.
- Creates the workbook and a `Delivery Report` worksheet when the report does not already exist.
- Formats the header row in bold and adjusts column widths, capped at 50 characters.
- If a PDF is not available but a PO number was found elsewhere, writes a report row containing the PO number and leaves the other fields blank.

### SAP GUI integration

- Connects through the SAP GUI COM scripting interface using `win32com.client`.
- Requires SAP GUI to be open, logged in, and configured with scripting enabled.
- Opens transaction `ME23N`.
- Tries the packaged SAP screen identifiers to locate the PO number field and enters the detected purchase-order number.
- Uses the SAP Generic Object Services menu to create a file attachment.
- Attaches the original Outlook `.msg` file, not the extracted PDF.
- Waits for the SAP attachment window and fills its folder and filename fields before confirming the upload.

### Logging and file movement

- Writes dated log files named like `Logs/log_YYYYMMDD.txt`.
- Moves successfully completed messages to `Processed`.
- Moves messages with missing PO numbers or processing errors to `Failed`.
- Adds a timestamp to the destination name if a move would overwrite an existing file.

## Technology stack

| Technology | Purpose |
| --- | --- |
| Python | Application and automation logic. The supplied executable embeds Python 3.14. |
| Tkinter | Windows desktop user interface. |
| `tkinterdnd2` | Drag-and-drop support for `.msg` files. |
| `extract-msg` | Reading Outlook `.msg` messages and attachments. The bundled build reports version 0.55.0. |
| `pdfplumber` | Extracting text from PDF attachments. |
| `openpyxl` | Creating, updating, formatting, and reading the Excel report. |
| `pywin32` / `win32com.client` | SAP GUI COM scripting integration. |
| PyInstaller | Packaging the Python application as `SAP_Email_Attacher.exe`. |
| SAP GUI for Windows | Target business system and GUI scripting endpoint. |
| Microsoft Excel | Used by the `Open in Excel` action to open the generated workbook. |

The application uses standard-library modules such as `pathlib`, `shutil`, `threading`, `datetime`, `re`, and `time`. No pandas dataframes, database, web API, SMTP client, or browser automation library was observed in the application modules.

## Application workflow

1. The application determines its base directory. In a frozen executable it uses the executable's parent directory; in source mode it uses the source file's parent directory.
2. It creates the runtime folders `Input`, `Temp`, `Processed`, `Failed`, `Logs`, and `Reports` when needed.
3. The user adds `.msg` files with the file picker or by dragging them into the window. Files are copied into `Input`; non-`.msg` files are skipped.
4. The user selects `Run Now`. The GUI disables the button, clears the progress panel, and runs the batch pipeline on a background thread.
5. The pipeline finds `Input/*.msg`. If there are no messages, it stops without connecting to SAP.
6. The pipeline connects once to the currently open SAP GUI session.
7. Each message is skipped if the same filename is already present in `Processed` or `Failed`.
8. `MsgReader` opens the message, reads its subject/sender/body, and extracts its attachments to `Temp`.
9. The pipeline searches for a PO number in the filename, subject, and body. If necessary, it extracts text from PDF attachments in attachment order until a match is found. The last PDF encountered by the attachment loop is retained for report-field extraction.
10. If no PO number is found, the message is logged as failed and moved to `Failed`.
11. If a PO number is found, the application extracts PDF fields when a PDF is available, appends a row to `Delivery_Report.xlsx`, and reports whether the row came from PDF data or contains only the PO number.
12. The SAP client opens the PO in `ME23N` and attaches the original `.msg` file through SAP Generic Object Services.
13. The message reader is closed, the message is logged as successful, and the `.msg` file is moved to `Processed`.
14. After the batch, the GUI re-enables `Run Now`, refreshes the input list, and refreshes the report tab.

## Installation

### Environment requirements

This application is Windows-specific because it uses SAP GUI COM scripting, `os.startfile`, and Windows Outlook `.msg` files.

You need:

- Windows.
- Python with Tkinter/Tcl-Tk installed when running from source. The supplied build embeds Python 3.14, but source compatibility should be verified against the original development environment.
- SAP GUI for Windows, already open and logged in.
- SAP GUI scripting enabled and permitted by the organisation's SAP policy.
- Microsoft Excel for the `Open in Excel` convenience action.
- Access to the relevant SAP purchase orders and authorization to create SAP attachments.

### Run from source

After placing the original source files in the recommended layout below:

```powershell
git clone <your-repository-url>
cd sap-purchase-order-email-attacher

py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

# Start the desktop application
python gui_app.py
```

The packaged modules also include a direct batch entry point:

```powershell
python main.py
```

The GUI entry point is the recommended user-facing start command. The direct batch entry point uses the same pipeline but without the desktop window.

### Run the supplied executable

The supplied `SAP_Email_Attacher.exe` is a PyInstaller build and already contains the Python runtime and bundled packages. Place it in a dedicated working folder and start it on a prepared Windows workstation. The runtime folders are created next to the executable.

Do not assume that the executable removes the need for SAP GUI, SAP scripting permissions, or Microsoft Excel. Those are external Windows/application requirements.

## Project structure

The following layout reflects the embedded module names and the runtime folders created by the application. Runtime folders should normally be ignored by Git and created on the user's machine.

```text
sap-purchase-order-email-attacher/
├── README.md
├── requirements.txt
├── .gitignore
├── gui_app.py                 # Tkinter/TkinterDnD2 desktop interface
├── main.py                    # Batch processing workflow
├── config.py                  # Base directory and runtime folder paths
├── core/
│   ├── __init__.py
│   ├── msg_reader.py          # Outlook MSG reading and attachment extraction
│   ├── po_extractor.py        # PO number and PDF text extraction
│   ├── report_builder.py      # Delivery_Report.xlsx creation/update
│   └── sap_client.py          # SAP GUI scripting and attachment upload
├── Input/                     # Runtime input .msg files; do not commit business data
├── Temp/                      # Runtime extracted attachments; do not commit
├── Processed/                # Successfully processed messages; do not commit
├── Failed/                   # Failed messages; do not commit
├── Logs/                     # Dated processing logs; do not commit
└── Reports/                 # Delivery_Report.xlsx; do not commit business data
```

`__init__.py`, `.gitignore`, tests, and build files were not present in the supplied executable as user-authored source artifacts; they are recommended repository files to add when reconstructing or publishing the source project.

## Inputs and outputs

| Type | Location/format | Details |
| --- | --- | --- |
| Input | Outlook `.msg` files | Added through the GUI or placed in `Input`. |
| Input | Email metadata | Filename, subject, body, and attachment names are searched for a PO number. |
| Input | PDF attachment(s) | Text is extracted when the PO is not found earlier and/or when report fields are needed. The current loop retains the last PDF encountered for report extraction. |
| Output | `Reports/Delivery_Report.xlsx` | Structured delivery report with eight columns. |
| Output | `Logs/log_YYYYMMDD.txt` | Timestamped processing status, PO number, success, skip, and failure messages. |
| Output | `Temp/` | Extracted attachment files used during processing. |
| Output | `Processed/` | Successfully handled `.msg` files. |
| Output | `Failed/` | Files that could not be completed, including missing-PO cases. |
| SAP output | SAP PO attachment | The original `.msg` file is attached to the matching SAP PO. |

The report is not a general-purpose data-analysis model. It is a structured operational delivery report populated from fixed text patterns in the PDF retained by the attachment loop.

## Business benefits

Based on the verified workflow and the author's operational context, the application can:

- Reduce manual movement of Outlook emails and attachments between Outlook, local folders, Excel, and SAP.
- Standardize purchase-order lookup using a consistent PO pattern and search priority.
- Reduce manual entry of delivery details into a workbook when the PDF text matches the expected formats.
- Preserve the original email as a SAP attachment for operational traceability.
- Provide separate success and failure queues for follow-up.
- Avoid reprocessing a message that is already present in `Processed` or `Failed`.
- Produce a dated log that helps users trace what happened to each file.

The reported time improvement is approximately one minute manually versus under 10 seconds through the application, subject to SAP response time, email size, workstation performance, and successful PO matching.

No independent time-savings, accuracy, throughput, or ROI measurements were run during this review.

## Security and privacy review

### Findings from the supplied executable

- The application modules contain no obvious hard-coded passwords, API keys, SMTP credentials, or usernames.
- SAP access is obtained from the currently open SAP GUI session through COM; no SAP password is stored in the recovered application logic.
- No company name, internal URL, or external API endpoint was observed in the recovered application modules.
- This is not a guarantee that the original source, Git history, build scripts, or runtime data are free of secrets. Review those separately before publishing.

### Files and data that must not be committed

Do not publish real business data, including:

- Outlook `.msg` files.
- PDF delivery documents.
- Extracted files in `Temp`.
- Messages in `Processed` and `Failed`.
- `Reports/Delivery_Report.xlsx`.
- `Logs/`, because logs may contain filenames, PO numbers, error details, and other operational information.
- Any SAP screenshots, exports, configuration files, or local path listings.
- `.env` files, credentials, certificates, tokens, test exports, and private keys if they exist in the original project.

The `.exe` itself embeds the Python bytecode and third-party libraries. If the source is proprietary, decide whether to keep the binary private, publish it only as a versioned release asset, or omit it from the repository. Do not use GitHub as a storage location for confidential `.msg`/PDF samples.

## Suggested `.gitignore` entries

```gitignore
.venv/
__pycache__/
*.py[cod]
*.pyd
build/
dist/
*.spec
*.exe

Input/*
Temp/*
Processed/*
Failed/*
Logs/*
Reports/*

!.gitkeep
.env
*.pem
*.key
```

If the executable is intentionally distributed, remove `*.exe` from the release workflow only and publish it as a reviewed release artifact rather than mixing it with source and business data.

## Limitations observed in the current implementation

These are current behaviors, not recommendations:

- PDF attachments are checked in order for PO fallback, while the last PDF encountered is retained for report-field extraction.
- PO matching is based on a fixed `45` plus eight digits pattern.
- PDF field extraction depends on fixed labels and text layouts; scanned/image-only PDFs may not yield text.
- The packaged code currently assigns the matched Customer Product value to both `Customer Product` and `Customer Product Code`.
- SAP screen identifiers and transaction paths are tied to the current SAP GUI scripting layout.
- The GUI copies files by filename into `Input`; a same-name copy can overwrite the existing input file before processing.
- Temporary extracted attachments are not visibly cleaned up by the recovered workflow.
- Source-level tests, CI configuration, structured error types, and a dependency lockfile were not present in the supplied artifact.

## Suggested future improvements

The following are not implemented features; they are optional next steps for the repository:

- Move SAP transaction/screen identifiers and folder settings into validated configuration or environment variables.
- Add unit tests for PO matching, attachment filename sanitization, PDF field extraction, duplicate handling, and report updates.
- Add integration tests with mocked SAP COM objects and representative non-confidential fixtures.
- Add explicit file-size/type validation and malware-scanning guidance for incoming `.msg` and PDF files.
- Add a per-file processing manifest or stable identifier instead of filename-only idempotency.
- Clean temporary attachments after successful processing, with a configurable retention policy for troubleshooting.
- Add structured logging with log levels and redaction rules for filenames, PO numbers, and message content.
- Add a dependency lockfile and a repeatable PyInstaller build workflow.
- Improve PDF extraction for scanned documents with a separately approved OCR workflow.
- Add a clear retry queue for recoverable SAP or file-system failures.

## Validation status

Static analysis of the supplied PyInstaller bundle confirmed the module names, imports, user-facing controls, workflow strings, PO pattern, report headers, runtime folders, SAP transaction path, and attachment behavior described above. The application was not executed, SAP was not connected, and no test email or PDF was processed during this review.
