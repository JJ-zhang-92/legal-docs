---
name: legal-docs
description: Use when the user asks to process, draft, review, or analyze Chinese legal documents, contracts, litigation filings, or perform case/statute research. Handles contract clause extraction, risk identification, document template generation, case law summarization, citation formatting, and OCR for scanned documents. Supports Word (.docx) and PDF files. Requires python-docx, pdfplumber, openpyxl, pytesseract, pdf2image.
---

# Legal Document Processing Skill

Process, draft, review, and analyze Chinese legal documents. Includes built-in modules for document templates, statute lookup, and OCR. Always write Python scripts to temp files and execute them via `bash` rather than using inline `-c` for non-trivial tasks.

## CRITICAL: File Naming — NEVER Overwrite

**This is the highest-priority rule. Must be followed without exception.**

When generating any output file (.docx, .xlsx, .pdf, .txt, etc.), you MUST:

1. **Check if the target path already exists** using `Test-Path -LiteralPath` before saving
2. **If the file exists, append a version suffix**: `_v2`, `_v3`, `_v4`... increment until an unused filename is found
3. **Never overwrite an existing file** — every output is a new file
4. **Pattern**: `filename.docx` → `filename_v2.docx` → `filename_v3.docx` ...

```powershell
# PowerShell implementation
$base = "C:\work\报告.docx"
if (Test-Path -LiteralPath $base) {
    $v = 2
    while (Test-Path -LiteralPath ($new = $base -replace '\.docx$', "_v$v.docx")) { $v++ }
    $base = $new
}
# Save to $base
```

```python
# Python implementation
import os
base = r'C:\path\to\output.docx'
if os.path.exists(base):
    v = 2
    name, ext = os.path.splitext(base)
    while os.path.exists(f"{name}_v{v}{ext}"):
        v += 1
    base = f"{name}_v{v}{ext}"
# Save to base
```

**Example behavior:**
- First run: `合同审查报告.docx` (created)
- Second run: `合同审查报告_v2.docx` (created, original untouched)
- Third run: `合同审查报告_v3.docx` (created, all previous versions preserved)

This ensures every generated document is traceable and no work is ever lost.

## CRITICAL: Output Version — Original vs De-identified

**When generating a docx that contains content from user documents, you MUST ask the user which version to output before saving.**

### How it works

```
AI reads deidentified_text ← full semantic structure, zero PII
        │
        ▼
AI decides: "use paragraphs 3, 5, 8-12, legal basis from para 15"
        │
        ▼
AI writes bash script ← the script does the disk I/O
   - Opens original file (docx/python-docx OR raw text)
   - Reads paragraphs by index [3, 5, 8, 9, 10, 11, 12, 15]
   - Writes selected original text into docx
   - stdout prints ONLY: "Done: output.docx (X paragraphs, Y chars)"
        │          ↑
        ▼          ╹
   Original text: disk → disk only, never through stdout → never reaches API

The paragraph_map from deidentify() bridges this:
  deid_text at index N → original[start:end] on disk → extracted by bash
```

### Rule

1. **Ask the user**: "输出原文版还是脱敏版？"
   - **原文版** → bash script reads original file directly and inserts selected paragraphs.
     The original text travels: disk → disk. API never sees it.
   - **脱敏版** → Uses already-deidentified text from the pipeline.

2. **Bash script MUST NOT print original content to stdout.**
   Only print: `"Done: filename.docx (N paragraphs, M chars)"`

3. **Use paragraph_map for precision extraction.**
   `deidentify()` returns `paragraph_map: [{index, start, end, deid_text, original_text}]`.
   AI decides which paragraphs to include by reading `deid_text`.
   Bash script extracts `original_text` by index from the original file.

4. **Save with clear suffix, never overwrite.**
   - 原文版 → `报告.docx`
   - 脱敏版 → `报告_脱敏版.docx`
   - Existing files → `_v2`, `_v3` etc.

## CRITICAL: Confidential Document Handling — Minimize Upload

**User's legal documents (contracts, complaints, evidence, etc.) are confidential.** Every file you read enters the AI provider's servers. Follow these rules strictly:

### 1. Get confirmation before reading user documents
- When the user asks to process a document, first explain what will be read and how
- Ask for explicit confirmation before reading, especially for full-text reads
- NEVER read user documents automatically — always prompt first

### 2. Minimize what you read
- **Default: extract structure only** — read just enough to get headings, parties, clause titles. Do NOT read the full text unless the user explicitly asks
- For contract review: extract clause titles + party info only, ask user which clauses they want analyzed in detail
- For complaint/defense generation: extract case number, court, parties, dispute type — ask user for the facts they want to include rather than reading the full document
- Full document read is ONLY for situations the user explicitly requests it

### 3. Warn after processing confidential content
- After processing user documents, remind: "已处理涉密文件内容已进入当前对话。如需控制扩散风险，建议后续清空对话或使用临时会话。"
- Remind the user that `/share` would expose the conversation content

### 4. Separate tools from case files
- The skill directory (`.opencode/skills/legal-docs/`) contains only tools and reference data — no case files
- User's actual legal documents should live outside the skill directory

### Quick reference for AI: when to read vs when to ask

| User request | Action |
|-------------|--------|
| "审查这份合同" | Extract parties + clause headings only, show overview, ask which clauses to review |
| "起草起诉状" | Ask for case facts verbally first, only read existing docs if user insists |
| "对比两份合同第X条" | Read only the specified clauses, not the full documents |
| "全文阅读这份判决书" | Confirm "将读取全文（XX页），确认？" then read |
| "查一下民法典第584条" | Query statutes.py locally via bash, no document upload needed |

## MANDATORY: De-Identification Pipeline Before API Access

**Every user document (contracts, complaints, judgments, etc.) MUST go through this pipeline before any content enters the AI context. The original document stays local — only the abstract reaches the API.**

### Pipeline

```
User document (.docx/.pdf/.txt)
        │
        ▼
  deidentifier.py    ← LOCAL ONLY: PII detection + redaction
        │              Replaces names, IDs, amounts, dates, addresses
        │              with role-aware placeholders
        ▼
  abstractor.py      ← LOCAL ONLY: structure extraction
        │              Detects doc type, extracts clause headings,
        │              dispute type, legal basis, party roles
        ▼
  JSON abstract       ← UPLOAD ONLY THIS to AI API
  (~300-800 chars)      Contains: doc_type, parties(roles only),
                         clause_headings, dispute_types, legal_basis
                         NO real names, NO amounts, NO dates, NO IDs
```

### Implementation (always write to temp script, execute via bash)

```python
import sys
sys.path.insert(0, r'PATH\TO\legal-docs')
from deidentifier import deidentify_file
from abstractor import generate_abstract
import json

# Step 1+2: De-identify + generate abstract (local only)
deid = deidentify_file(r'user_document.docx')
abstract = generate_abstract(deid)

# Step 3: Save abstract locally (for reference)
with open(r'output_abstract.json', 'w', encoding='utf-8') as f:
    json.dump(abstract, f, ensure_ascii=False, indent=2)

# Step 4: Read ONLY the abstract into AI context
# (NEVER read the original or deidentified full text)
```

### What the abstract contains (safe to upload)

| Field | Example | Notes |
|-------|---------|-------|
| doc_type | "民事起诉状" | Detected automatically |
| parties | {"原告":"原告","被告":"被告"} | Roles only, no real names |
| clause_headings | ["第一条 出资","第二条 分工"] | Structure only |
| dispute_types | ["侵害商标权","不正当竞争"] | Keywords extracted |
| legal_basis | ["商标法第57条","反不正当竞争法第6条"] | Public law references |
| claims | ["判令被告停止侵权..."] | Sanitized claim text |

### What is REDACTED (never uploaded)

- Company names, personal names → replaced with [甲方]/[原告] etc.
- ID numbers, credit codes, bank accounts → [身份证号已脱敏] etc.
- Phone numbers, emails → [手机号已脱敏]
- Monetary amounts → [人民币金额已脱敏]
- Dates → [完整日期已脱敏]
- Physical addresses → [详细地址已脱敏]
- Case numbers → [案号已脱敏]

### Rules for AI behavior

1. **Run pipeline first, talk later**: Before analyzing any user document, execute the deidentify→abstract pipeline via bash. Read only the JSON abstract into context.
2. **Ask for specific clauses**: If you need more detail, ask the user "需要查看第X条的具体内容吗？" — then re-run the pipeline on just that clause.
3. **Never shortcut**: Do NOT read the original .docx/.pdf directly into context, even if the user asks. Explain: "为保护信息安全，我先本地生成脱敏摘要，再基于摘要进行分析。"
4. **Abstract is the canonical reference**: All AI analysis must be based on the abstract. If the abstract lacks detail, request a targeted clause extraction.

### Folder Input (Directory Processing)

When the user provides a folder path (e.g., "审查 C:\案卷\"):

```
Step 1: directory_scanner.scan_directory() → LOCAL bash only
        → Every supported file runs FULL deidentifier→abstractor pipeline
        → Each file gets a label: doc_type + parties + key content hint
        → Output: numbered file index with pipeline-extracted labels (~30-80 bytes/file)
        → ZERO original content enters AI context

Step 2: Show index to user → "目录含N个文件，处理哪些？"
        → User selects specific files by number or description

Step 3: For selected files — use the already-generated abstracts from Step 1
        → Abstracts are cached locally from the scan phase
        → Only selected abstracts enter AI context for analysis
```

**Implementation:**
```python
from directory_scanner import scan_directory, format_index_for_ai

index = scan_directory(user_folder_path)       # LOCAL: pipeline per file
print(format_index_for_ai(index))              # Safe AI context: labels only
# AI shows numbered list → user picks → AI uses cached abstracts
```

**Rules:**
- EVERY supported file (.docx/.pdf/.txt) is pipeline-processed — no filename shortcuts
- NEVER auto-read all files without user selection
- Labels are extracted from content, NOT from filenames
- Unsupported types (images, spreadsheets) are flagged but not auto-processed

## Prerequisites

```powershell
pip install python-docx pdfplumber openpyxl pytesseract Pillow pdf2image
```

For OCR: install Tesseract-OCR with Chinese language pack.
- Windows: https://github.com/UB-Mannheim/tesseract/wiki
- Check: `python statutes\ocr_helper.py --check`

## Supporting Modules

All modules are located in the same directory as this SKILL.md. Import paths assume scripts are written to `$env:TEMP`:

| Module | Purpose |
|--------|---------|
| `directory_scanner.py` | **New** — Folder scanner: metadata → pipeline labels → index for AI |
| `deidentifier.py` | **PII脱敏引擎** — role-aware entity replacement, 11 entity types, local only, batch support |
| `abstractor.py` | **结构化摘要生成器** — doc type detection, clause extraction, legal basis extraction, batch support |
| `templates\document_templates.py` | Generate formatted .docx: 起诉状, 答辩状, 法律意见书, 律师函, 委托代理合同, 授权委托书 |
| `rag_client.py` | **New** — Local RAG client for law knowledge base, fallback to statutes.py |
| `statutes.py` | Search/lookup 民法典 (全7编), 公司法, 劳动合同法, 民事诉讼法, 刑法, 江浙沪地方法规 |
| `ocr_helper.py` | OCR scanned PDFs/images with Chinese support, preprocessing, batch processing |
| `bid_utils.py` | **New** — Bid requirement extraction, completeness check, package integrity, report generation |
| `contract_utils.py` | PDF page rendering, contract front+signature page collage for bid appendices |
| `verify_cross_reference.py` | Cross-reference validation between bid body and qualification appendix |
| `ppt-rules.md` | PPT generation technical specifications and best practices |

---

## 1. Contract Review & Analysis

> **Pipeline note:** The code below illustrates the analysis logic. When processing actual user documents, first run the de-identification pipeline (Section "MANDATORY") to generate the abstract, then work from the abstract. The read-and-analyze code here is the SECOND step after the pipeline.

### Read contracts
```python
import os
from docx import Document
import pdfplumber

for root, dirs, files in os.walk(base_directory):
    for f in files:
        if f.endswith('.docx') and not f.startswith('~$'):
            doc = Document(os.path.join(root, f))
            text = '\n'.join(p.text for p in doc.paragraphs)
        elif f.endswith('.pdf'):
            with pdfplumber.open(os.path.join(root, f)) as pdf:
                text = ''.join(p.extract_text() or '' for p in pdf.pages)
```

### Extract key clauses
```python
clause_patterns = {
    '违约': r'(第[零一二三四五六七八九十百千]+条[^。]*?违约[^。]*。)',
    '争议解决': r'(第[零一二三四五六七八九十百千]+条[^。]*?(?:仲裁|诉讼|管辖)[^。]*。)',
    '保密': r'(第[零一二三四五六七八九十百千]+条[^。]*?(?:保密|商业秘密)[^。]*。)',
    '价款': r'(第[零一二三四五六七八九十百千]+条[^。]*?(?:价款|价格|金额|费用)[^。]*。)',
    '知识产权': r'(第[零一二三四五六七八九十百千]+条[^。]*?(?:知识产权|专利|商标|著作权)[^。]*。)',
    '不可抗力': r'(第[零一二三四五六七八九十百千]+条[^。]*?不可抗力[^。]*。)',
}

def extract_clauses(text, patterns):
    results = {}
    for name, pat in patterns.items():
        matches = re.findall(pat, text)
        results[name] = matches if matches else ['[未找到相关条款]']
    return results
```

### Risk clause detection
```python
risk_patterns = {
    '单方权利过大': r'(?:有权单方|自行决定|无需通知|无需经[^；。]{1,10}同意)',
    '责任限制不清': r'(?:间接损失|利润损失[^；。]*不[^；。]*承担|不[^；。]*超过[^；。]*元)',
    '违约金过高': r'违约金[^；。]*每[日天][^；。]*(?:[1-9]|10)%',
    '质保期过短': r'质保(?:期|期限)[^；。]*(\d+)[个]?(?:天|日|月)',
}

def detect_risks(text, patterns):
    risks = []
    for risk_name, pat in patterns.items():
        for m in re.finditer(pat, text):
            risks.append({
                'risk_type': risk_name,
                'location': f'位置 {m.start()}-{m.end()}',
                'text': m.group()[:200],
                'severity': 'high' if risk_name in ('单方权利过大', '违约金过高') else 'medium'
            })
    return risks
```

### Output review report
Generate Markdown report with sections:
1. 合同基本信息 (parties, amount, term)
2. 条款清单 (clause inventory table)
3. 风险条款 (risk table with severity)
4. 缺失条款 (recommended missing clauses)
5. 修改建议 (specific revision suggestions with 民法典 article references from statutes.py)

---

## 2. Legal Document Drafting

Use `templates\document_templates.py` — it provides full .docx generators with proper Chinese legal formatting.

### Available templates
```python
import sys
sys.path.insert(0, r'PATH\TO\legal-docs\templates')
from document_templates import TEMPLATE_REGISTRY, generate_complaint, generate_answer, generate_legal_opinion, generate_demand_letter, generate_agency_agreement, generate_power_of_attorney
```

### Generate 起诉状
```python
output = r'C:\work\起诉状.docx'
generate_complaint(output,
    plaintiff_name='张三',
    plaintiff_gender='男',
    plaintiff_dob='1985.06.15',
    plaintiff_ethnic='汉族',
    plaintiff_address='北京市朝阳区XX街XX号',
    plaintiff_id='110101198506150000',
    plaintiff_phone='13800000000',
    defendant_name='XX房地产有限公司',
    defendant_is_company=True,
    defendant_legal_rep='李四',
    defendant_address='北京市海淀区XX路XX号',
    defendant_credit_code='91110000XXXXXXXXXX',
    court_name='北京市海淀区人民法院',
    claims=[
        '判令被告支付拖欠的工程款人民币500,000元及利息',
        '判令被告承担本案全部诉讼费用',
    ],
    facts='2024年3月1日，原告与被告签订《XX项目施工合同》，约定...',
    evidence_list=[
        '《XX项目施工合同》原件一份',
        '工程结算单一份',
        '银行转账记录三份',
        '微信聊天记录截图',
    ],
)
```

### Generate other documents
```python
# 答辩状
generate_answer(output,
    respondent_name='XX公司',
    case_number='（2024）京0105民初12345号',
    court_name='北京市朝阳区人民法院',
    claimant_name='张三',
    cause_of_action='买卖合同纠纷',
    response_points=['原告诉讼请求无事实依据', '原告主张的违约金数额过高'],
    facts='...')

# 法律意见书
generate_legal_opinion(output,
    client_name='XX投资有限公司',
    matter='关于XX收购项目的法律尽职调查',
    background='...',
    legal_analysis='...',
    risks=['目标公司存在未决诉讼', '知识产权归属不明'],
    conclusion='建议在交割前完成以下事项...',
    law_firm='XX律师事务所',
    lawyer_name='王律师')

# 律师函
generate_demand_letter(output,
    sender_name='XX科技有限公司',
    recipient_name='YY股份有限公司',
    fact_statement='...',
    legal_basis='根据《民法典》第577条...',
    demands=['立即停止侵权行为', '三日内支付拖欠款项'],
    deadline_days=7,
    law_firm='XX律师事务所',
    lawyer_name='赵律师')

# 委托代理合同
generate_agency_agreement(output,
    client_name='李四', law_firm='XX律师事务所',
    lawyer_name='刘律师', matter='XX合同纠纷案',
    fee_amount='50000', fee_type='固定收费')

# 授权委托书
generate_power_of_attorney(output,
    principal_name='王五', attorney_name='刘律师',
    law_firm='XX律师事务所', case_matter='XX纠纷',
    opponent_name='赵六', scope='特别授权')
```

### Document formatting conventions
- Title: SimHei (黑体), 二号(22pt), bold, centered
- Section headers: SimHei, 三号(16pt), bold
- Body text: SimSun (宋体), 四号(14pt)
- Line spacing: 1.5x
- Margins: top/bottom 2.54cm, left 3.17cm (binding), right 2.54cm
- First-line indent: 0.74cm (two Chinese characters)

---

## 3. Statute & Case Law Research

**Primary: `rag_client.py`** — queries a local law-knowledge-base RAG service at `http://localhost:8720`. Returns full statute text, not just summaries. When the RAG service is unavailable, automatically falls back to `statutes.py` keyword search.

```python
from rag_client import query
result = query("劳动合同试用期最长多久", top=3)
# Returns full statute text from RAG, or keyword summaries from statutes.py fallback
```

**Secondary: `statutes.py`** — covers 民法典 all 7 volumes, 公司法, 劳动合同法, 民事诉讼法, 刑法 (economic crimes). Used as fallback when RAG is offline.

### Search by keywords
```python
import sys
sys.path.insert(0, r'PATH\TO\legal-docs')
from statutes import search_statutes, get_article_text, get_chapter_outline

# Search across all statutes
results = search_statutes(['违约金', '违约责任', '解除合同'])
for r in results:
    print(f"{r['statute']} | {r['chapter']} | {r['article']}")
    print(f"  {r['summary'][:100]}...")
```

### Look up specific article
```python
print(get_article_text('民法典', 584))  # → 第584条：损害赔偿范围
print(get_article_text('民法典', '第143条'))  # → 民事法律行为有效条件
print(get_article_text('公司法', 182))  # → 董监高禁止利用公司机会
print(get_article_text('劳动合同法', 47))  # → 经济补偿计算标准
```

### Get statute outline
```python
print(get_chapter_outline('民法典'))
print(get_chapter_outline('公司法'))
```

### Extract case information from judgment texts
```python
import re

info = {}
info['案号'] = re.search(r'([（(][12]\d{3}[）)][\u4e00-\u9fa5]?\w?[字第初终再裁\d]+号)', text)
info['法院'] = re.search(r'([\u4e00-\u9fa5]+(?:人民法院|仲裁委员会))', text)
info['裁判结果'] = re.findall(r'(?:判决|裁定)如下[：:]\s*(.*?)(?:本案|如不|审\s*判|本判决)', text, re.S)
info['法律依据'] = re.findall(r'(?:《[^》]*?》第\s*\d+\s*条)', text)
```

### Statute coverage reference

#### National statutes

| Statute | Articles | Key areas |
|---------|----------|-----------|
| 民法典·总则 | 1-204 | 民事主体、法律行为、代理、诉讼时效 |
| 民法典·物权 | 205-462 | 所有权、用益物权、担保物权 |
| 民法典·合同 | 463-978 | 合同成立/效力/履行/违约、买卖合同、借款、保证、租赁、委托 |
| 民法典·人格权 | 989-1039 | 名誉、隐私、个人信息 |
| 民法典·婚姻家庭 | 1040-1118 | 结婚、离婚、夫妻财产、子女抚养 |
| 民法典·继承 | 1119-1163 | 法定继承、遗嘱、遗产处理 |
| 民法典·侵权 | 1164-1258 | 过错/无过错责任、损害赔偿、产品/医疗/交通/环境责任 |
| 公司法 (2024.7.1) | 1-260 | 公司设立、股东权利、董监高义务、增减资、解散清算 |
| 劳动合同法 | — | 合同订立/解除、试用期、经济补偿/赔偿金 |
| 民事诉讼法 | — | 管辖、证据、一审/二审/再审、保全、执行 |
| 刑法 | — | 职务侵占、挪用资金、合同诈骗、侵犯商业秘密等经济犯罪 |

#### 江浙沪地方法规

```python
from statutes import list_local_regulations, search_local_regulations

# List all local laws
print(list_local_regulations())

# Filter by region
print(list_local_regulations('浙江省'))

# Search only local regulations
results = search_local_regulations(['工伤', '工资'], region='浙江省')
for r in results:
    print(f"[{r['region']}] {r['statute']} | {r['article']}: {r['summary']}")

# Search all statutes including local (use search_statutes)
all_results = search_statutes(['数字经济', '数据'])
```

**浙江省（8部）：**
| 法规 | 施行日期 | 关键条文 |
|------|----------|----------|
| 民营企业发展促进条例 | 2020.2.1 | 平等准入、禁止歧视性条件、不得拖欠款项(8条) |
| 数字经济促进条例 | 2021.3.1 | 数据要素市场、公共数据开放、产业数字化(5条) |
| 电子商务条例 | 2022.3.1 | 直播营销管理、信用评价、知识产权保护(5条) |
| 知识产权保护和促进条例 | 2023.1.1 | 惩罚性赔偿、海外维权、技术调查官(5条) |
| 工伤保险条例 | 2018.1.1 | 工伤认定、停工留薪期、未参保处理(5条) |
| 劳动人事争议调解仲裁条例 | 2016.1.1 | 先行调解、终局裁决标准、45日审限(5条) |
| 人身损害赔偿统一标准通知 | 2020.4.1 | 同命同价、统一适用城镇标准(3条) |
| 企业工资支付管理办法 | 2017.5.1 | 欠薪垫付、应急周转金制度(5条) |

**江苏省（6部）：**
| 法规 | 施行日期 | 关键条文 |
|------|----------|----------|
| 劳动合同条例 | 2013.5.1 | 试用期规定、高温保护、合并分立处理(5条) |
| 工资支付条例 | 2005.1.1(2019修正) | 工资计算起点、延期支付限制、欠薪垫付(5条) |
| 物业管理条例 | 2001.3.1(2020修正) | 业主大会、物业服务收费、维修资金(4条) |
| 人身损害赔偿标准通知 | 2020.3.20 | 同命同价(3条) |
| 社会保险费征缴条例 | 2014.1.1 | 参保时限、滞纳金(3条) |
| 数字经济促进条例 | 2022.8.1 | 产业数字化、数据要素市场(3条) |

**上海市（5部）：**
| 法规 | 施行日期 | 关键条文 |
|------|----------|----------|
| 劳动合同条例 | 2002.5.1(2015修正) | 合同必备条款、裁员程序、非全日制(4条) |
| 工伤保险实施办法 | 2013.1.1 | 工伤认定、停工留薪待遇、补助金标准(5条) |
| 数据条例 | 2022.1.1 | 数据权益保护、公共数据分类开放、浦东先行先试(5条) |
| 知识产权保护条例 | 2021.3.1 | 惩罚性赔偿、展会保护、快速维权(4条) |
| 优化营商环境条例 | 2020.4.10(2023修正) | 告知承诺制、政府采购公平、破产退出(4条) |

**长三角区域协作（4项）：**
| 文件 | 时间 | 核心内容 |
|------|------|----------|
| 生态绿色一体化发展示范区总体方案 | 2019.11 | 企业自由迁移、要素流动一体化(4条) |
| 知识产权保护合作协议 | 2019 | 跨区域执法联动、信息共享(3条) |
| 法院司法协作框架协议 | 2019 | 跨域立案、异地执行查控、统一裁判标准(4条) |
| 跨区域行政争议实质性化解机制 | 2022 | 异地管辖、提级管辖(2条) |

---

## 4. Multi-Contract Comparison

> Pipeline first: run deidentification on all contracts, compare at the abstract level. Full text comparison only with explicit user confirmation. The code below is template logic.

```python
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

def compare_contracts(doc1_text, doc2_text, label1='合同A', label2='合同B'):
    paras1 = [p for p in doc1_text.split('\n') if p.strip()]
    paras2 = [p for p in doc2_text.split('\n') if p.strip()]
    matcher = difflib.SequenceMatcher(None, paras1, paras2)

    wb = Workbook()
    ws = wb.active
    ws.title = '合同对比'
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

    ws.append(['变更类型', f'{label1}（原文）', f'{label2}（对比文本）', '差异说明', '风险等级'])

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for k in range(i1, i2):
                ws.append(['相同', paras1[k], paras2[j1 + k - i1], '', ''])
        elif tag == 'replace':
            for k in range(max(i2 - i1, j2 - j1)):
                old = paras1[i1 + k] if i1 + k < i2 else '[已删除]'
                new = paras2[j1 + k] if j1 + k < j2 else '[新增]'
                row = ws.append(['变更', old, new, '条款内容被修改', '中'])
                for cell in ws[ws.max_row]:
                    cell.fill = yellow_fill
        elif tag == 'delete':
            for k in range(i1, i2):
                row = ws.append(['删除', paras1[k], '', '条款在对比文本中不存在', '高'])
                for cell in ws[ws.max_row]:
                    cell.fill = red_fill
        elif tag == 'insert':
            for k in range(j1, j2):
                row = ws.append(['新增', '', paras2[k], '新增条款', '高'])
                for cell in ws[ws.max_row]:
                    cell.fill = green_fill

    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 10
    ws.freeze_panes = 'A2'
    return wb
```

---

## 5. OCR for Scanned Documents

Use `ocr_helper.py` for scanned PDFs and images. **OCR output is auto-deidentified** — `_deidentify_if_available()` runs on all text returns. No raw OCR text enters AI context.

### Smart OCR (auto-detects digital vs scanned)
```python
import sys
sys.path.insert(0, r'PATH\TO\legal-docs')
from ocr_helper import smart_ocr, ocr_pdf, ocr_image, extract_with_preprocessing

# Auto-detect: extracts text from digital PDF, OCR for scanned pages
result = smart_ocr(r'合同扫描件.pdf', lang='chi_sim+eng')
print(result['full_text'])
print(f"Pages: {result['page_count']}, OCR'd pages: {result.get('ocr_pages', [])}")

# Force full OCR
result = ocr_pdf(r'扫描文件.pdf', lang='chi_sim+eng', dpi=300)
```

### Single image OCR
```python
text = ocr_image(r'判决书照片.jpg', lang='chi_sim+eng')
```

### Enhanced OCR for low-quality scans
```python
text = extract_with_preprocessing(r'模糊扫描件.png',
    output_dir=r'C:\temp\preprocessed')
```

### Batch OCR
```python
from ocr_helper import batch_ocr_directory

results = batch_ocr_directory(r'D:\案卷材料', output_file='ocr_results.txt')
```

### OCR troubleshooting
- Check installation: `python ocr_helper.py --check`
- Missing Chinese: download chi_sim.traineddata to Tesseract tessdata directory
- Low quality: increase DPI to 400, use `extract_with_preprocessing`
- Mixed content: `smart_ocr` handles digital+scanned PDFs automatically

---

## 6. Legal Date & Deadline Calculator

```python
from datetime import datetime, timedelta

legal_periods = {
    '上诉期-民事': 15, '上诉期-刑事': 10, '上诉期-行政': 15,
    '举证期限': 30, '答辩期': 15, '行政复议': 60, '行政诉讼': 180,
    '保全续期': 180, '申请执行': 720,  # days
}

def calculate_deadline(start_date, period_days, holidays=None):
    end_date = start_date
    days_added = 0
    holiday_set = set(holidays or [])
    while days_added < period_days:
        end_date += timedelta(days=1)
        if end_date.weekday() < 5:
            if (end_date.month, end_date.day) not in holiday_set:
                days_added += 1
    return end_date
```

## 7. Bid Document Processing (投标文件处理)

Bid documents are legal documents — they carry contractual obligations, qualification declarations, and legal liability. The skill provides dedicated tooling for the full bid lifecycle.

### Supporting modules

| Module | Purpose |
|--------|---------|
| `bid_utils.py` | **New** — Requirement extraction, completeness check, package integrity, report generation |
| `verify_cross_reference.py` | Cross-reference validation between bid body and qualification appendix |
| `contract_utils.py` | PDF page rendering, contract front+signature page collage for bid appendices |

### 7a. Extract tender requirements
```python
from bid_utils import extract_requirements

reqs = extract_requirements(r'招标文件.docx')
# Returns structured dict:
#   project_name, bid_deadline, qualification_requirements,
#   technical_requirements, commercial_requirements,
#   required_documents, evaluation_criteria, contact_info
```

### 7b. Generate completeness checklist
```python
from bid_utils import generate_checklist

checklist = generate_checklist(reqs)
# Returns flat list: [(category, item, status), ...]
# Categories: 基础信息, 资格要求, 技术要求, 商务要求, 提交文件
```

### 7c. Check bid vs requirements
```python
from bid_utils import compare_response, generate_bid_report

# Full comparison: tender requirements vs bid response
result = compare_response(r'投标书.docx', r'招标文件.docx')
print(result['summary'])  # {total_checks, found, missing, completeness_rate}

# Or: comprehensive report with cross-reference + package check
report = generate_bid_report(r'投标书.docx', r'招标文件.docx', r'投标文件包/')
```

### 7d. Package integrity check
```python
from bid_utils import check_package_integrity

pkg = check_package_integrity(r'D:\投标文件包\')
# Checks for: 投标函, 报价单, 资格证明, 授权书, 技术方案,
#  业绩证明, 商务偏离表, 售后服务承诺, 中小企业声明函
print('OK' if pkg['integrity_ok'] else f'Missing: {pkg["missing"]}')
```

### 7e. Cross-reference verification
```python
from verify_cross_reference import verify, print_report

issues = verify(r'投标书_v2.docx', r'资格证明文件_v2.docx')
print_report(issues)
# Checks that every contract/entity referenced in the bid body
# has a corresponding entry in the qualification appendix
```

### 7f. Contract image collage for bid appendices
```python
from contract_utils import make_contract_collage, batch_collage

# Single contract: front page + signature page side-by-side
make_contract_collage(r'合同.pdf', dpi=200)

# Batch: process all PDFs in a directory
batch_collage(r'D:\合同文件\', output_dir=r'D:\投标附录\')
```

### Typical bid workflow

```
招标文件.docx ──→ bid_utils.extract_requirements() ──→ 需求清单
                    │
投标书.docx ──────→ bid_utils.compare_response() ──→ 完整性报告
                    │
资格证明文件.docx ─→ verify_cross_reference.verify() ──→ 交叉引用报告
                    │
合同文件.pdf ─────→ contract_utils.make_contract_collage() ──→ 附录图片
                    │
投标文件包/ ──────→ bid_utils.check_package_integrity() ──→ 包完整性检查
                    │
                    ▼
            bid_utils.generate_bid_report() ──→ 综合审查报告（Markdown）
```

## File Handling Guidelines

1. **Always use raw strings** (`r'path'`) for Windows file paths
2. **Write scripts to temp**: `$env:TEMP\legal_script.py`, execute with `python "$env:TEMP\legal_script.py"`
3. **Encoding**: Always use `encoding='utf-8'` for Chinese text
4. **Clean up temp files** after use
5. **Verify output** with `Test-Path` and `Get-Item`
6. **Never hard-code Chinese paths** -- use `os.walk()` to discover files
7. **Module imports**: Copy needed `.py` files to `$env:TEMP` alongside the script, or use `sys.path.insert()` to point to the skill directory

## Output Formatting

- **Tables** for comparison data (clauses, risks, parties)
- **Bullet points** for lists (evidence, procedural steps)
- **Blockquotes** for legal citations
- **Bold** for critical risk items
- When citing statutes, use full reference: `《民法典》第577条`

## Limitations & Disclaimers

1. **NOT a substitute for professional legal advice** -- always recommend review by a qualified attorney
2. **Statute references are approximate** -- must be verified against official current versions
3. **This skill does NOT provide legal advice** -- it assists with document processing, formatting, and statute navigation
4. **Draft documents** must be reviewed and finalized by licensed legal professionals
