# legal-docs

OpenCode skill for Chinese legal document processing — contracts, litigation, bid documents, and OCR.

Privacy-first: all sensitive document processing happens locally. Only de-identified structured abstracts reach the AI API.

## Quick Start

```powershell
pip install python-docx pdfplumber openpyxl pytesseract Pillow pdf2image pymupdf
```

Place this directory in your OpenCode skills folder:

```
.opencode/skills/legal-docs/
```

## Modules

| Module | Purpose |
|--------|---------|
| `deidentifier.py` | PII redaction engine — 11 entity types, role-aware replacement |
| `abstractor.py` | Structured abstract generator — 11 document types, clause extraction |
| `directory_scanner.py` | Folder scanner — batch pipeline with file index labels |
| `statutes.py` | Statute database — Civil Code + Company Law + local regulations |
| `templates/document_templates.py` | 6 legal document generators (complaint, defense, opinion...) |
| `ocr_helper.py` | OCR for scanned PDFs/images — Chinese + English |
| `bid_utils.py` | Bid document tools — requirement extraction, completeness check |
| `contract_utils.py` | PDF rendering, contract collage for bid appendices |
| `verify_cross_reference.py` | Cross-reference validation (bid body vs qualification appendix) |
| `ppt-rules.md` | PPT generation best practices |

## Workflow

```
User document → deidentifier (local) → abstractor (local) → JSON abstract → API
                                                                  ↑
                                                         ~300-800 chars only
                                                       No names/amounts/dates
```

All document content stays on your machine. Only structured abstracts reach the AI.

## Security

- **Zero PII upload**: names, IDs, phone numbers, addresses, amounts, dates are redacted locally
- **Never overwrite**: output files auto-versioned (`_v2`, `_v3`...)
- **Folder-safe**: directory scans process labels only, not content
- **Output-aware**: user confirms original vs de-identified version before docx write

## License

MIT
