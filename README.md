# SAP Purchase Order Email Attacher

Windows-focused procurement automation for processing Outlook order-confirmation messages, extracting purchase-order information, updating an Excel delivery report, and attaching the original email to the matching SAP purchase order through SAP GUI scripting.

> Repository status: This public repository currently contains documentation and dependency metadata only. It does not currently include the standalone Python source, packaged executable, tests, fixtures, or a release artifact. A fresh clone is therefore not runnable. The capability description below is based on static analysis of an external packaged build and must be validated against the original source before production use.

## Business problem

Procurement teams may need to move order-confirmation information between Outlook, PDF attachments, Excel delivery reports, and SAP. A typical manual process is to find the purchase-order number, open the purchase order in SAP, attach the original email for traceability, enter delivery details into Excel, and record failures for follow-up.

The external packaged build reviewed for this project was designed to standardize that workflow. The current GitHub repository does not contain the implementation needed to run or independently verify it.

## What is currently in this repository

| Item | Status |
| --- | --- |
| README.md | Project description, workflow notes, limitations, and publication guidance |
| requirements.txt | Dependencies recovered from an external packaged build |
| .gitignore | Rules for runtime data and local secrets |
| GITHUB_METADATA.md | Suggested repository metadata and publication checklist |
| Python source | Not included |
| Packaged executable | Not included and no GitHub release is available |
| Tests and CI | Not included |
| Screenshots or sample data | Not included |

## Observed workflow in the external build

The following behavior was observed through static analysis of a packaged build. It is not a claim that the current repository contains these modules or that the behavior has been independently re-tested.

1. The application creates runtime folders named Input, Temp, Processed, Failed, Logs, and Reports.
2. A user adds Outlook MSG files through a desktop interface or places them in Input.
3. The workflow searches the filename, subject, body, and PDF text for a 10-digit purchase-order value beginning with 45.
4. PDF text can be used to populate a delivery report when the document matches the expected text layout.
5. The original MSG file is attached to the matching SAP purchase order through SAP GUI scripting and transaction ME23N.
6. Completed files are moved to Processed; files that cannot be completed are moved to Failed.
7. Dated log files are written under Logs.

## Observed capabilities

- Tkinter desktop interface with file picker, Run Now action, input count, progress output, report view, and drag-and-drop support through tkinterdnd2.
- Outlook MSG reading through extract-msg, including subject, sender, body, and attachments.
- Purchase-order matching in filename, subject, body, and PDF text. The observed pattern is 45 followed by eight digits.
- PDF text extraction through pdfplumber and Excel report creation through openpyxl.
- SAP GUI COM scripting through pywin32. The external build uses ME23N and Generic Object Services for the attachment step.
- Separate success and failure queues, dated logs, duplicate checks, and filename collision handling.

## Technology stack observed

| Technology | Purpose |
| --- | --- |
| Python | Application and automation logic |
| Tkinter | Windows desktop interface |
| tkinterdnd2 | Drag-and-drop support |
| extract-msg | Outlook MSG parsing |
| pdfplumber | PDF text extraction |
| openpyxl | Excel report generation and updates |
| pywin32 | SAP GUI COM integration |
| PyInstaller | Packaging the external executable |
| SAP GUI for Windows | Target business system |
| Microsoft Excel | Opening and reviewing the generated report |

## Installation and current run status

This repository does not currently provide a supported run command because the source files and packaged executable are not included.

The dependency file can be installed for inspection with:

~~~powershell
git clone https://github.com/prasadpai567-debug/sap-purchase-order-email-attacher.git
cd sap-purchase-order-email-attacher
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
~~~

Installing dependencies alone does not produce an application. Do not run this workflow against production SAP data until the original source, test evidence, and deployment process have been reviewed.

If the source is later approved for publication, this README should be updated with the verified entry point, supported Python version, configuration steps, build instructions, and a tested example.

## External prerequisites for the observed build

When and if the implementation is published, the observed workflow would require Windows, SAP GUI for Windows, an authorized logged-in SAP GUI session, SAP GUI scripting enabled according to organizational policy, Microsoft Excel, and permission to create SAP attachments.

## Inputs and outputs observed

| Type | Location or format | Notes |
| --- | --- | --- |
| Input | Outlook MSG files | Added through the interface or placed in Input |
| Input | Email metadata | Filename, subject, body, and attachment names can be searched |
| Input | PDF attachments | Extraction depends on the expected document layout |
| Output | Reports/Delivery_Report.xlsx | Structured delivery report in the external build |
| Output | Logs/log_YYYYMMDD.txt | Dated processing status in the external build |
| Output | Processed and Failed | Operational file queues in the external build |
| SAP output | SAP purchase-order attachment | Original MSG file in the external build |

## Error handling and logging observed

The external build documented separate success and failure queues, dated log files, duplicate checks, and filename collision handling. These behaviors cannot be reviewed at source level because the source is not present in this repository.

Before publishing or using the implementation, verify exception handling around MSG parsing, PDF extraction, Excel writes, file moves, and SAP COM operations. Also verify retry behavior, temporary-file cleanup, log redaction, and partial-failure reporting.

## Business benefits and evidence

If validated against the original source and a controlled test set, this project could demonstrate reduced manual movement of order-confirmation emails between Outlook, Excel, local folders, and SAP; more consistent purchase-order matching; reduced manual entry of delivery details; improved SAP traceability; and clearer failure follow-up.

The previous README reported approximately one minute manually versus under ten seconds through the automation. This is user-reported, was not benchmarked in this repository, and should not be presented as a measured KPI until repeatable tests document the sample size, SAP response time, workstation, and success rate.

## Security and privacy

No supplier files, MSG files, PDFs, generated reports, logs, credentials, or private keys are currently committed. The ignore rules exclude the main runtime folders, environment files, private keys, executables, and generated output.

Do not publish real Outlook MSG files, PDF delivery documents, customer or supplier data, SAP exports, generated reports, operational logs, screenshots containing internal information, credentials, tokens, certificates, private keys, or local configuration files.

The source and any external executable must still be scanned before publication. Dependency versions should be pinned and locked after the original source is restored and tested.

## Limitations observed in the external build

- Purchase-order matching uses a fixed 45 plus eight digits pattern.
- PDF extraction depends on fixed labels and text layouts; scanned PDFs may not produce usable text.
- SAP screen identifiers depend on the SAP GUI layout.
- Filename-only duplicate handling may not identify the same message under a renamed file.
- Temporary attachment retention requires explicit verification.
- Source-level tests, CI, structured error types, and a dependency lockfile were not available for review.

## Demonstration materials

No screenshots or sanitized sample files are currently included. After the implementation is approved for publication, add only non-confidential material such as a redacted desktop screenshot, a synthetic MSG-to-report example, a short test-environment demonstration, or a sample report with fictitious values.

## Recommended next steps

1. Decide whether the original Python source can be published safely.
2. If approved, add the source with a reproducible entry point and configuration example.
3. Add unit tests for PO matching, filename sanitization, PDF parsing, report updates, and duplicate handling.
4. Add mocked integration tests for SAP and Outlook boundaries.
5. Add CI for tests, dependency review, and secret scanning.
6. Pin dependencies and document the supported Python version.
7. Add sanitized screenshots and a versioned release only after validation.
8. Add a license after selecting the intended reuse terms.

## Recruiter-friendly project summary

Procurement automation project that streamlines order-confirmation handling by reading Outlook MSG files, identifying purchase orders, extracting delivery data from PDFs, updating an Excel report, and attaching the original email to the corresponding SAP purchase order through SAP GUI scripting. The project demonstrates Python automation, document processing, Excel integration, SAP GUI scripting, operational logging, and exception workflow design.

The current GitHub repository contains documentation and dependency metadata only; the runnable source is not currently published.

## Validation status

This repository was reviewed at the GitHub main branch. The current tree contains four documentation or configuration files and no executable application or tests. No SAP session, Outlook mailbox, Excel workbook, MSG file, PDF, or production system was accessed or modified during this review.
