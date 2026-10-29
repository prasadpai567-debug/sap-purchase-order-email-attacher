# GitHub repository metadata

## Repository description

Windows Python automation tool that reads Outlook order-confirmation and PO-related `.msg` emails, identifies the purchase order, updates an Excel delivery report when applicable, and attaches the original email to the matching SAP purchase order through SAP GUI scripting.

## Suggested repository name

`sap-purchase-order-email-attacher`

## Suggested GitHub topics

```text
python
python-automation
sap
sap-gui
sap-gui-scripting
pywin32
win32com
outlook
msg-files
pdfplumber
excel-automation
openpyxl
procurement
purchase-orders
document-processing
tkinter
tkinterdnd2
desktop-application
rpa
```

## Recommended repository structure

```text
sap-purchase-order-email-attacher/
├── README.md
├── requirements.txt
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
├── tests/                  # Recommended; add non-confidential fixtures only
├── Input/.gitkeep          # Runtime only
├── Temp/.gitkeep           # Runtime only
├── Processed/.gitkeep      # Runtime only
├── Failed/.gitkeep         # Runtime only
├── Logs/.gitkeep           # Runtime only
└── Reports/.gitkeep        # Runtime only
```

Do not commit real `.msg` files, PDFs, generated Excel reports, logs, SAP exports, or credentials. The supplied executable was reviewed as a packaged artifact; standalone Python source was not included in the submitted files.
