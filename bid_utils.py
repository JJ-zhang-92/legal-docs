"""
Bid document processing utilities.
Extracts requirements from tender documents, checks completeness,
generates checklists, and compares bid responses against requirements.

Usage:
    from bid_utils import extract_requirements, check_completeness, generate_checklist, compare_response
    reqs = extract_requirements('招标文件.docx')
    checklist = generate_checklist(reqs)
    issues = check_completeness('投标书.docx', checklist)
"""

import re, os
from docx import Document
from datetime import datetime


def _read_text(filepath):
    """Read full text from .docx file. Auto-deidentifies when deidentifier
    module is available (inside skill context). Standalone imports skip this."""
    import sys, os
    # Ensure skill directory is on sys.path for cross-module import
    skill_dir = os.path.dirname(os.path.abspath(__file__))
    if skill_dir not in sys.path:
        sys.path.insert(0, skill_dir)

    try:
        doc = Document(filepath)
        text = '\n'.join(p.text for p in doc.paragraphs)
    except (FileNotFoundError, OSError, ValueError) as e:
        raise FileNotFoundError(f'Cannot read file: {filepath} — {e}') from e

    try:
        from deidentifier import deidentify
        return deidentify(text)['deidentified_text']
    except ImportError:
        raise ImportError(
            'The deidentifier module is required for bid document processing. '
            'Please ensure bid_utils is running from within the legal-docs skill directory.'
        ) from None


# ═══════════════════════════════════════════════════════════════
# 1. TENDER REQUIREMENT EXTRACTION
# ═══════════════════════════════════════════════════════════════

def extract_requirements(filepath):
    """
    Extract structured requirements from a tender document (招标文件).

    Args:
        filepath: Path to .docx tender document

    Returns:
        dict with keys:
            'project_name': Project name
            'bid_deadline': Submission deadline
            'qualification_requirements': List of qualification requirements
            'technical_requirements': List of technical requirements
            'commercial_requirements': List of commercial requirements
            'required_documents': List of required submission documents
            'evaluation_criteria': List of evaluation criteria
            'contact_info': Tender organizer contact info
    """
    text = _read_text(filepath)
    return _parse_requirements(text)


def _parse_requirements(text):
    """Parse tender requirements from text."""
    result = {
        'project_name': None,
        'bid_deadline': None,
        'qualification_requirements': [],
        'technical_requirements': [],
        'commercial_requirements': [],
        'required_documents': [],
        'evaluation_criteria': [],
        'contact_info': {},
    }

    # ── Project name ──
    m = re.search(r'(?:项目名称|采购项目|招标项目)[：:]\s*(.{5,50})', text)
    if m:
        result['project_name'] = m.group(1).strip()

    # ── Bid deadline ──
    m = re.search(r'(?:投标截止|递交截止|开标)[时日期间].{0,10}[：:]\s*(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日]?)', text)
    if m:
        result['bid_deadline'] = m.group(1).strip()

    # ── Qualification requirements ──
    qual_section = _extract_section(text, ['投标人资格', '资格要求', '合格投标人', '资格条件'])
    if qual_section:
        # Split by numbered items
        items = re.split(r'(?:\n\s*)?\d+[\.\、\)）]\s*', qual_section)
        items = [i.strip() for i in items if len(i.strip()) > 5]
        result['qualification_requirements'] = items

    # ── Technical requirements ──
    tech_section = _extract_section(text, ['技术需求', '技术要求', '技术规格', '技术参数', '服务要求'])
    if tech_section:
        items = re.split(r'(?:\n\s*)?\d+[\.\、\)）]\s*', tech_section)
        items = [i.strip() for i in items if len(i.strip()) > 5]
        result['technical_requirements'] = items

    # ── Commercial requirements ──
    comm_section = _extract_section(text, ['商务要求', '商务条款', '报价要求', '付款方式', '合同条款'])
    if comm_section:
        items = re.split(r'(?:\n\s*)?\d+[\.\、\)）]\s*', comm_section)
        items = [i.strip() for i in items if len(i.strip()) > 5]
        result['commercial_requirements'] = items

    # ── Required documents ──
    doc_section = _extract_section(text, [
        '投标文件组成', '投标文件内容', '需提交的文件', '应提交的材料',
        '投标文件要求', '响应文件组成', '响应文件内容'
    ])
    if doc_section:
        items = re.split(r'(?:\n\s*)?\d+[\.\、\)）]\s*', doc_section)
        result['required_documents'] = [i.strip()[:80] for i in items if len(i.strip()) > 3]

    # ── Evaluation criteria ──
    eval_section = _extract_section(text, ['评审标准', '评审办法', '评标办法', '评分标准', '评审因素'])
    if eval_section:
        items = re.split(r'(?:\n\s*)?\d+[\.\、\)）]\s*', eval_section)
        result['evaluation_criteria'] = [i.strip()[:100] for i in items if len(i.strip()) > 5]

    # ── Contact info ──
    for field, pattern in [
        ('采购人', r'采购人[：:]?\s*(.{3,30})'),
        ('代理机构', r'(?:招标|采购)代理机构[：:]?\s*(.{3,30})'),
        ('联系人', r'联系\s*人[：:]?\s*(.{2,10})'),
        ('电话', r'(?:联系\s*)?电\s*话[：:]?\s*(\d[\d\- ]{6,20})'),
        ('地址', r'地\s*址[：:]?\s*(.{5,50})'),
    ]:
        m = re.search(pattern, text)
        if m:
            result['contact_info'][field] = m.group(1).strip()

    return result


def _extract_section(text, keywords):
    """Extract the section of text containing any of the given keywords.
    Keywords must appear at a line/section boundary to avoid false matches
    (e.g., '技术' should match '技术要求' but not '信息技术')."""
    for kw in keywords:
        # Match kw at start of line or after heading markers (numbers, bullets)
        m = re.search(rf'(?:^|\n)\s*(?:[\d一二三四五六七八九十]+[\.\、\)）]?\s*)?{re.escape(kw)}', text)
        if m:
            idx = m.start()
            # Find the section start (look backward for a heading)
            start = text.rfind('\n', 0, idx)
            start = text.rfind('\n', 0, start) if start > 50 else max(0, idx - 50)

            # Find section end (next major heading or 2000 chars)
            end = idx + 2000
            next_heading = re.search(r'\n\s*[一二三四五六七八九十]、|\n\s*第[一二三四五六七八九十]章|\n\s*\d+[\.\、]', text[idx:end])
            if next_heading:
                end = idx + next_heading.start()
            return text[idx:end].strip()[:3000]

    return None


# ═══════════════════════════════════════════════════════════════
# 2. COMPLETENESS CHECKLIST GENERATION
# ═══════════════════════════════════════════════════════════════

def generate_checklist(requirements):
    """
    Generate a completeness checklist from extracted requirements.

    Args:
        requirements: dict from extract_requirements()

    Returns:
        List of checklist items, each a tuple: (category, item, status)
    """
    checklist = []

    # Basic info
    checklist.append(('基础信息', '项目名称已确认', 'pending'))
    checklist.append(('基础信息', '投标截止时间已记录', 'pending'))
    checklist.append(('基础信息', '开标时间/地点已确认', 'pending'))

    # Qualification
    for i, req in enumerate(requirements.get('qualification_requirements', []), 1):
        checklist.append(('资格要求', f'资格{i}: {req[:60]}', 'pending'))

    # Technical
    for i, req in enumerate(requirements.get('technical_requirements', []), 1):
        checklist.append(('技术要求', f'技术{i}: {req[:60]}', 'pending'))

    # Commercial
    for i, req in enumerate(requirements.get('commercial_requirements', []), 1):
        checklist.append(('商务要求', f'商务{i}: {req[:60]}', 'pending'))

    # Required documents
    for i, doc_item in enumerate(requirements.get('required_documents', []), 1):
        checklist.append(('提交文件', f'文件{i}: {doc_item[:60]}', 'pending'))

    return checklist


# ═══════════════════════════════════════════════════════════════
# 3. COMPLETENESS CHECK (Bid vs Requirements)
# ═══════════════════════════════════════════════════════════════

def check_completeness(bid_doc_path, requirements_or_checklist):
    """
    Check if the bid document addresses all requirements.

    Args:
        bid_doc_path: Path to bid response .docx
        requirements_or_checklist: Either dict from extract_requirements()
                                   or list from generate_checklist()

    Returns:
        List of issues: [(category, item, status, detail)]
    """
    bid_text = _read_text(bid_doc_path)

    # Convert requirements to checklist if needed
    if isinstance(requirements_or_checklist, dict):
        checklist = generate_checklist(requirements_or_checklist)
    else:
        checklist = requirements_or_checklist

    issues = []
    for category, item, _ in checklist:
        # Check if the requirement topic appears in the bid text
        # Extract key terms from the item (first 10+ meaningful chars)
        key_terms = re.sub(r'^[^:：]+[：:]\s*', '', item)[:20].strip()

        if key_terms and len(key_terms) >= 3:
            found = key_terms in bid_text
            status = 'found' if found else 'missing'
            issues.append({
                'category': category,
                'item': item,
                'status': status,
                'detail': '已在投标书中找到相关描述' if found else '可能遗漏'
            })

    return issues


# ═══════════════════════════════════════════════════════════════
# 4. BID RESPONSE vs REQUIREMENT COMPARISON
# ═══════════════════════════════════════════════════════════════

def compare_response(bid_doc_path, tender_doc_path):
    """
    Full pipeline: extract requirements → check bid completeness.

    Args:
        bid_doc_path: Path to bid response .docx
        tender_doc_path: Path to tender document .docx

    Returns:
        dict with 'requirements', 'issues', 'summary'
    """
    reqs = extract_requirements(tender_doc_path)
    issues = check_completeness(bid_doc_path, reqs)

    total = len(issues)
    found = sum(1 for i in issues if i['status'] == 'found')
    missing = sum(1 for i in issues if i['status'] == 'missing')

    return {
        'requirements': reqs,
        'issues': issues,
        'summary': {
            'total_checks': total,
            'found': found,
            'missing': missing,
            'completeness_rate': f'{round(found*100/total)}%' if total > 0 else 'N/A',
        }
    }


# ═══════════════════════════════════════════════════════════════
# 5. BID PACKAGE INTEGRITY CHECK
# ═══════════════════════════════════════════════════════════════

def check_package_integrity(bid_dir):
    """
    Check completeness of a bid package directory.

    Args:
        bid_dir: Directory containing bid files

    Returns:
        dict with 'files', 'missing_expected', 'size_report'
    """
    expected_files = [
        ('投标函', ['投标函', '投标书', '响应函']),
        ('报价单', ['报价', '开标一览表', '价格表']),
        ('资格证明', ['资格', '资质', '营业执照']),
        ('法定代表人授权书', ['授权书', '授权委托']),
        ('技术方案', ['技术方案', '技术响应', '技术偏离']),
        ('业绩证明', ['业绩', '合同', '案例']),
        ('商务偏离表', ['商务偏离', '商务条款偏离']),
        ('售后服务承诺', ['售后服务', '服务承诺']),
        ('中小企业声明函', ['中小企业声明', '小微企业']),
    ]

    files_in_dir = []
    for root, dirs, files in os.walk(bid_dir):
        for f in files:
            if not f.startswith('~$'):
                files_in_dir.append(f)

    all_files = ' '.join(files_in_dir).lower()
    found = []
    missing = []

    for label, keywords in expected_files:
        if any(kw.lower() in all_files for kw in keywords):
            found.append((label, '已找到'))
        else:
            missing.append((label, '缺失'))

    # Size report
    size_report = {}
    for f in files_in_dir:
        fpath = os.path.join(bid_dir, f) if os.path.isabs(bid_dir) else os.path.join(os.path.abspath(bid_dir), f)
        # Walk through subdirs
        for root, dirs, filenames in os.walk(bid_dir):
            if f in filenames:
                fpath = os.path.join(root, f)
                break
        try:
            size_mb = os.path.getsize(fpath) / (1024 * 1024)
            if size_mb > 0.1:  # Only report files > 100KB
                size_report[f] = round(size_mb, 2)
        except:
            pass

    return {
        'files': files_in_dir,
        'found': found,
        'missing': missing,
        'size_report': size_report,
        'integrity_ok': len(missing) == 0,
    }


# ═══════════════════════════════════════════════════════════════
# 6. REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════

def generate_bid_report(bid_doc_path, tender_doc_path=None, bid_dir=None):
    """
    Generate a comprehensive bid package report in Markdown format.

    Args:
        bid_doc_path: Path to the main bid document .docx
        tender_doc_path: Optional path to tender document for requirement comparison
        bid_dir: Optional directory path for package integrity check

    Returns:
        Formatted Markdown report string
    """
    lines = [
        f'# 投标文件审查报告',
        f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M")}',
        '',
    ]

    lines.append(f'投标文件：{os.path.basename(bid_doc_path)}')
    lines.append('')

    # Requirement comparison
    if tender_doc_path:
        lines.append('## 一、招标要求响应检查')
        result = compare_response(bid_doc_path, tender_doc_path)
        lines.append(f'- 检查项：{result["summary"]["total_checks"]}')
        lines.append(f'- 已响应：{result["summary"]["found"]}')
        lines.append(f'- 可能遗漏：{result["summary"]["missing"]}')
        lines.append(f'- 完整度：{result["summary"]["completeness_rate"]}')
        lines.append('')

        if result['issues']:
            lines.append('### 遗漏提醒')
            for issue in result['issues']:
                if issue['status'] == 'missing':
                    lines.append(f'- [{issue["category"]}] {issue["item"]}')

        lines.append('')

    # Cross-reference check
    lines.append('## 二、交叉引用校验')
    lines.append('> 使用 verify_cross_reference.py 检查投标书与资格证明文件之间的一致性')
    lines.append('> 执行方式: `from verify_cross_reference import verify, print_report`')
    lines.append('')

    # Package integrity
    if bid_dir:
        lines.append('## 三、投标包完整性')
        pkg = check_package_integrity(bid_dir)
        lines.append(f'- 包完整性：{"通过" if pkg["integrity_ok"] else "不完整"}')
        lines.append(f'- 文件数量：{len(pkg["files"])}')
        lines.append('')

        if pkg['found']:
            lines.append('### 已包含文件')
            for label, status in pkg['found']:
                lines.append(f'- [{status}] {label}')

        if pkg['missing']:
            lines.append('### 缺失文件')
            for label, status in pkg['missing']:
                lines.append(f'- [{status}] {label}')  

        if pkg['size_report']:
            lines.append('### 文件大小')
            for fname, size in sorted(pkg['size_report'].items(), key=lambda x: x[1], reverse=True):
                lines.append(f'- {fname}: {size} MB')

    lines.append('')
    lines.append('---')
    lines.append('*本报告由 bid_utils.py 自动生成，仅供参考。最终审核应由专业人员完成。*')

    return '\n'.join(lines)


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Usage:")
        print("  python bid_utils.py extract <招标文件.docx>")
        print("  python bid_utils.py checklist <招标文件.docx>")
        print("  python bid_utils.py verify <投标书.docx> <招标文件.docx>")
        print("  python bid_utils.py package <投标目录>")
        print("  python bid_utils.py report <投标书.docx> [招标文件.docx] [投标目录]")
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == 'extract':
        print("# Auto-deidentified output (safe for AI context)")
        reqs = extract_requirements(sys.argv[2])
        for k, v in reqs.items():
            if v:
                print(f"\n## {k}")
                if isinstance(v, list):
                    for i, item in enumerate(v, 1):
                        print(f"  {i}. {item}")
                else:
                    print(f"  {v}")

    elif cmd == 'checklist':
        reqs = extract_requirements(sys.argv[2])
        cl = generate_checklist(reqs)
        for cat, item, status in cl:
            print(f"[{status}] [{cat}] {item}")

    elif cmd == 'verify':
        result = compare_response(sys.argv[2], sys.argv[3])
        print(f"Completeness: {result['summary']['completeness_rate']}")
        print(f"Found: {result['summary']['found']}, Missing: {result['summary']['missing']}")

    elif cmd == 'package':
        pkg = check_package_integrity(sys.argv[2])
        print(f"Integrity: {'OK' if pkg['integrity_ok'] else 'INCOMPLETE'}")
        for label, status in pkg['missing']:
            print(f"  MISSING: {label}")

    elif cmd == 'report':
        tender = sys.argv[3] if len(sys.argv) > 3 else None
        bid_dir = sys.argv[4] if len(sys.argv) > 4 else None
        print(generate_bid_report(sys.argv[2], tender, bid_dir))
