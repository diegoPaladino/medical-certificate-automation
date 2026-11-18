# AI-Powered Medical Certificate Processing & Workflow Automation

A Python desktop application that helps school staff process medical certificates, validate student records, and prepare absence justifications in a school administration portal.

**Current version: 1.3.17.** Built for a Brazilian school workflow, with a Portuguese interface and Windows scanner integration.

## What it does

- Extracts information from document images and PDFs through the OpenAI API.
- Checks student names against local CSV or Excel records, including approximate name matching.
- Acquires documents from Windows WIA scanners.
- Prepares and compresses images for the portal's upload limit.
- Automates login, form filling, and document attachment with Playwright.
- Refreshes the student register daily, preserves the previous dataset when collection fails, and maintains browser sessions.
- Saves documents with standardized filenames and verifies copies before removing originals.

## Engineering focus

This project combines a Tkinter desktop interface, document processing, external API integration, browser automation, and local data synchronization. Background work uses threads and queues to coordinate with the interface. The workflow includes session recovery, temporary file cleanup, and failure handling.

## Run locally

Use Python 3.10 or later with Tkinter available. Windows is required for WIA scanner acquisition.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
$env:OPENAI_API_KEY = "your-api-key"
python main.py
```

Configure portal credentials in the application or through `GEMUL_USUARIO` and `GEMUL_SENHA`. Portal automation requires authorized access to the configured GEMUL instance. Supply your own student register or use the authenticated synchronization workflow. The repository does not include production records or credentials.

The default document model is configurable through `OPENAI_MODEL`. Document extraction sends selected document content to the configured OpenAI API; API access and usage charges are separate. Review extracted information before completing the administrative workflow.

PyMuPDF handles PDF rendering. The optional `pdf2image` fallback requires a separate installation of pdf2image and Poppler. WIA acquisition requires a compatible scanner and Windows drivers.

## Project files

| File | Purpose |
| --- | --- |
| `main.py` | Current desktop application |
| `student_sync.py` | Student register refresh and validation |
| `test_sessao_alunos.py` | Session and student synchronization regression tests |
| `RELEASES.md` | Release history in Portuguese |
| `LeitorAtestado*.py` | Earlier source snapshots |

Run the existing regression suite with `python -m unittest test_sessao_alunos.py`.

Local credentials, medical documents, student records, logs, and synchronization backups are excluded through `.gitignore`.
