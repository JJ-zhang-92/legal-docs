"""
Cross-reference verification between bid body DOCX and qualification appendix DOCX.
Checks that all contracts/entities referenced in the bid body have corresponding
entries in the appendix, and vice versa.

Usage:
    from verify_cross_reference import verify, print_report
    issues = verify('投标书.docx', '资格证明文件.docx')
    print_report(issues)
"""

import re, os, sys
from docx import Document

# Add skill directory for deidentifier import
_skill_dir = os.path.dirname(os.path.abspath(__file__))
if _skill_dir not in sys.path:
    sys.path.insert(0, _skill_dir)

# --- Generic entity patterns (no real client names) ---
# Users should customize these patterns for their specific bid documents.
ENTITY_PATTERNS = {
    '政府机构': [
        (r'(?:部|委|局|中心|党校)', '政府部门/机构'),
    ],
    '国有企业': [
        (r'(?:集团|总公司|有限公司|股份公司)', '国有企业'),
        (r'(?:保险|铁路|电力|石油|电信|能源)', '国央企/行业'),
    ],
    '教育医疗': [
        (r'(?:大学|学院|学校|医院|研究院)', '教育医疗机构'),
    ],
    '民营企业': [
        (r'(?:科技|网络|信息|建设|投资).*?(?:有限公司|股份)', '民营企业'),
    ],
}

def _read_docx(filepath):
    """Read docx with auto-deidentification when available."""
    try:
        doc = Document(filepath)
        text = '\n'.join(p.text for p in doc.paragraphs)
        try:
            from deidentifier import deidentify
            return deidentify(text)['deidentified_text']
        except ImportError:
            raise ImportError(
                'The deidentifier module is required for document processing. '
                'Please ensure verify_cross_reference is running from within '
                'the legal-docs skill directory.'
            ) from None
    except (FileNotFoundError, OSError, ValueError) as e:
        raise FileNotFoundError(f'Cannot read file: {filepath} — {e}') from e


def extract_entities(text, patterns):
    """Find all entities matching given patterns in text."""
    found = set()
    for pattern, label in patterns:
        if re.search(pattern, text):
            found.add(label)
    return found


def verify(bid_body_path, appendix_path):
    """Cross-reference bid body and appendix. Returns list of issues."""
    bid_text = _read_docx(bid_body_path)
    qual_text = _read_docx(appendix_path)

    issues = []
    for category, patterns in ENTITY_PATTERNS.items():
        for pattern, label in patterns:
            # Check if entity is mentioned in bid body
            in_bid = re.search(pattern, bid_text)
            in_qual = re.search(pattern, qual_text)

            if in_bid and not in_qual:
                issues.append({
                    'type': 'missing_in_appendix',
                    'category': category,
                    'label': label,
                    'detail': f'{label} 在投标书中出现但资格证明中未找到',
                    'pattern': pattern,
                })
            elif in_qual and not in_bid:
                issues.append({
                    'type': 'missing_in_bid_body',
                    'category': category,
                    'label': label,
                    'detail': f'{label} 在资格证明中出现但投标书正文中未引用',
                    'pattern': pattern,
                })

    return issues


def print_report(issues):
    """Print a formatted cross-reference report."""
    if not issues:
        print("No cross-reference issues found.")
        return

    missing_in_appendix = [i for i in issues if i['type'] == 'missing_in_appendix']
    missing_in_bid = [i for i in issues if i['type'] == 'missing_in_bid_body']

    print(f"Cross-reference report: {len(issues)} issues found")
    print(f"  Missing from appendix: {len(missing_in_appendix)}")
    print(f"  Not referenced in bid body: {len(missing_in_bid)}")
    print()

    if missing_in_appendix:
        print("=== ENTITIES IN BID BUT NOT IN APPENDIX ===")
        for issue in missing_in_appendix:
            print(f"  [{issue['category']}] {issue['label']}")
        print()

    if missing_in_bid:
        print("=== ENTITIES IN APPENDIX BUT NOT IN BID BODY ===")
        for issue in missing_in_bid:
            print(f"  [{issue['category']}] {issue['label']}")
