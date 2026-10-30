# GitHub Repository Metadata and Publication Notes

## Current repository status

- Repository: sap-purchase-order-email-attacher
- Visibility: Public
- Default branch: main
- Current contents: README.md, GITHUB_METADATA.md, requirements.txt, and .gitignore
- Python source: Not currently included
- Packaged executable: Not currently included
- Tests, CI, screenshots, and releases: Not currently included

The public repository is currently documentation-only. The workflow description is based on static analysis of an external packaged build and must be validated against the original source before production use.

## Recommended GitHub description

Windows Python automation for Outlook MSG purchase-order emails, Excel delivery reporting, and SAP GUI attachment workflows.

## Recommended GitHub topics

Use only topics supported by the published implementation:

~~~text
python
python-automation
sap
sap-gui
sap-gui-scripting
pywin32
outlook
msg-files
pdfplumber
excel-automation
openpyxl
procurement
purchase-orders
document-processing
tkinter
desktop-application
rpa
~~~

Do not add topics for features that are only planned or described from an unavailable build.

## Recommended structure when source publication is approved

~~~text
sap-purchase-order-email-attacher/
├── README.md
├── requirements.txt or pyproject.toml
├── .gitignore
├── gui_app.py
├── main.py
├── config.py
├── core/
│   ├── __init__.py
│   ├── msg_reader.py
│   ├── po_extractor.py
│   ├── report_builder.py
│   └── sap_client.py
├── tests/
├── Input/.gitkeep
├── Temp/.gitkeep
├── Processed/.gitkeep
├── Failed/.gitkeep
├── Logs/.gitkeep
└── Reports/.gitkeep
~~~

The structure above is a recommendation, not a claim about files currently present.

## Publication checklist

- Confirm that the original source can be published without exposing company, supplier, customer, or SAP-sensitive information.
- Remove or sanitize MSG files, PDFs, Excel reports, logs, exports, screenshots, credentials, tokens, certificates, and private keys.
- Add a verified run command and supported Python version.
- Add tests for PO matching, PDF extraction, Excel report updates, duplicate handling, and filename sanitization.
- Add mocked SAP and Outlook integration tests where practical.
- Pin and lock dependencies after source compatibility is verified.
- Add CI for tests, dependency review, and secret scanning.
- Add sanitized screenshots or a synthetic demonstration.
- Add a versioned release only after the executable has been reviewed and tested.
- Select and add an appropriate license before inviting reuse.
