# legal-docs / 法律文书处理

OpenCode skill for Chinese legal document processing — contracts, litigation, bid documents, and OCR.

中文法律文书处理技能 — 合同审查 · 诉讼文书 · 投标文件 · OCR 识别。

Privacy-first: all sensitive document processing happens locally. Only de-identified structured abstracts reach the AI API.

隐私优先：所有涉密文件处理在本地完成。仅结构化脱敏摘要上传 AI API。

## Quick Start / 快速开始

```powershell
pip install python-docx pdfplumber openpyxl pytesseract Pillow pdf2image pymupdf
```

Place this directory in your OpenCode skills folder / 放入 OpenCode 技能目录：

```
.opencode/skills/legal-docs/
```

## Modules / 模块

| Module | Purpose / 用途 |
|--------|---------------|
| `deidentifier.py` | PII redaction — 11 entity types, role-aware / PII 脱敏 — 11类实体 |
| `abstractor.py` | Structured abstract — 11 document types / 结构化摘要 — 11种文书 |
| `directory_scanner.py` | Folder scanner — batch pipeline / 文件夹扫描 — 批量流水线 |
| `rag_client.py` | Local law RAG client — queries external knowledge base, falls back to statutes.py / 本地法规 RAG 客户端 — 查询外部知识库，不可用时回退 |
| `statutes.py` | Statute database — Civil Code + Company Law + local regulations / 法规数据库 — 民法典 + 公司法 + 地方法规 |
| `templates/document_templates.py` | 6 legal document generators / 6种文书生成器 |
| `ocr_helper.py` | OCR for scanned PDFs/images / 扫描件 OCR（中英文） |
| `bid_utils.py` | Bid document tools — requirement extraction, completeness check / 投标文件工具 |
| `contract_utils.py` | PDF rendering, contract collage / PDF 渲染、合同拼图 |
| `verify_cross_reference.py` | Cross-reference validation / 投标书交叉引用校验 |
| `ppt-rules.md` | PPT generation best practices / PPT 生成规范 |

## Law RAG Integration / 法规 RAG 集成

> The law knowledge base RAG service is **not included** in this repository. It resides in the author's separate project.
> 
> 法规知识库 RAG 服务不在本仓库中，位于作者的其他项目。

| Item / 项目 | Detail / 说明 |
|------------|---------------|
| 所在位置 / Location | 作者的独立项目 (author's separate project) |
| 端口 / Port | `localhost:8720` |
| 启动 / Start | `python law-rag/server.py` |
| 状态检查 / Check | `python rag_client.py --status` |
| 无 RAG 时 / Without RAG | 自动回退到 `statutes.py` 关键词搜索 / Auto-fallback to statutes.py keyword search |
| 查询 / Query | `python rag_client.py "劳动合同试用期" 3` |

## Workflow / 工作流

```
User document → deidentifier (local) → abstractor (local) → JSON abstract → API
                                                                   ↑
                                                          ~300-800 chars only
                                                        No names/amounts/dates
```

All document content stays on your machine. Only structured abstracts reach the AI.

所有文件内容留在本地。仅结构化摘要上传 AI。

## Security / 安全

- **Zero PII upload / 零实名上传**: names/IDs/phones/addresses/amounts/dates redacted locally
- **Never overwrite / 永不覆盖**: output files auto-versioned (`_v2`, `_v3`...)
- **Folder-safe / 文件夹安全**: directory scans process labels only, not content
- **Output-aware / 输出确认**: user confirms original vs de-identified version before docx write

## License / 许可证

MIT
