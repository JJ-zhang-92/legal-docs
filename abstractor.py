"""
Structured abstract generator for Chinese legal documents.
Processes de-identified text and produces a compact JSON abstract (~500 chars).
The abstract is the ONLY content uploaded to the AI API.

Zero external dependencies. Rule-based extraction — no LLM required.

Usage:
    from deidentifier import deidentify_file
    from abstractor import generate_abstract

    deid = deidentify_file('contract.docx')
    abstract = generate_abstract(deid['deidentified_text'], deid)
    print(json.dumps(abstract, ensure_ascii=False, indent=2))
"""

import re
import json
import os


# ═══════════════════════════════════════════════════════════════
# DOCUMENT TYPE DETECTION
# ═══════════════════════════════════════════════════════════════

DOC_TYPE_SIGNATURES = [
    ('民事起诉状', ['起诉状', '民事起诉状', '诉讼请求', '事实与理由', '此致']),
    ('民事答辩状', ['答辩状', '民事答辩状', '答辩意见', '答辩人']),
    ('民事判决书', ['判决书', '民事判决书', '本院认为', '判决如下', '如不服本判决']),
    ('民事裁定书', ['裁定书', '民事裁定书', '裁定如下']),
    ('三方合作协议', ['共同投资', '利润分配', '三方', '甲乙丙']),
    ('双方合同', ['甲方', '乙方', '合同', '协议']),
    ('律师函', ['律师函', '律师事务所', '致：', '委托']),
    ('法律意见书', ['法律意见书', '法律分析', '风险提示', '结论']),
    ('授权委托书', ['授权委托书', '受托人', '代理权限']),
    ('仲裁申请书', ['仲裁申请书', '仲裁请求', '仲裁委员会']),
    ('上诉状', ['上诉状', '上诉人', '原审判决']),
    ('证据清单', ['证据清单', '证据及证据来源', '证据名称']),
]

def detect_doc_type(text):
    """Detect legal document type by keyword signatures."""
    scores = {}
    for doc_type, keywords in DOC_TYPE_SIGNATURES:
        score = sum(1 for kw in keywords if kw in text)
        if score > 0:
            scores[doc_type] = score
    if not scores:
        return '通用法律文书'
    return max(scores, key=scores.get)


# ═══════════════════════════════════════════════════════════════
# STRUCTURE EXTRACTION
# ═══════════════════════════════════════════════════════════════

def _extract_headings(text):
    """Extract section/chapter/clause headings from the document."""
    headings = []

    # Match Chinese-style section headings: "第X条 XXXX" or "一、XXXX" or "（一）XXXX"
    NUM = r'[零一二三四五六七八九十百千]+'
    patterns = [
        rf'(第{NUM}章[：:\s]*[^\n]{{2,40}})',
        rf'(第{NUM}条[：:\s]*[^\n]{{2,60}})',
        rf'([一二三四五六七八九十]、[^\n]{{2,40}})',
        rf'(（[一二三四五六七八九十]）[^\n]{{2,40}})',
        rf'(第{NUM}节[：:\s]*[^\n]{{2,40}})',
    ]

    for pat in patterns:
        matches = re.findall(pat, text)
        for m in matches:
            clean = re.sub(r'^[一二三四五六七八九十]+[、）\.]?\s*', '', m.strip()).strip()
            if clean and len(clean) >= 3 and clean not in headings:
                headings.append(clean)

    return headings[:20]  # Cap at 20


def _extract_legal_basis(text):
    """Extract cited laws and articles."""
    cited = []

    # Match statute citations: 《XXX》第N条 (both Arabic and Chinese numerals)
    # Chinese numeral mapping for article numbers
    CN_NUM = r'(?:[零一二三四五六七八九十百千]+)'
    patterns = [
        # Arabic: 《民法典》第584条
        r'《([^》]{2,30})》\s*第\s*(\d+)\s*条',
        # Chinese: 《民法典》第五百八十四条
        rf'《([^》]{{2,30}})》\s*第\s*({CN_NUM})\s*条',
        # Fallback: law name only (民法典/刑法/公司法 et al.)
        r'《([^》]{2,30}[法法典例程])》',
    ]

    for pat in patterns:
        matches = re.findall(pat, text)
        for m in matches:
            if isinstance(m, tuple):
                citation = f'《{m[0]}》第{m[1]}条' if len(m) > 1 else f'《{m[0]}》'
            else:
                citation = m
            if citation not in cited:
                cited.append(citation)

    return cited[:15]


def _extract_dispute_type(text):
    """Extract dispute type / cause of action."""
    dispute_keywords = [
        '侵害商标权', '侵害专利权', '侵害著作权', '不正当竞争',
        '合同纠纷', '买卖合同纠纷', '借款合同纠纷', '租赁合同纠纷',
        '劳动纠纷', '劳动争议', '工伤', '经济补偿金',
        '机动车交通事故', '人身损害', '医疗损害',
        '离婚纠纷', '抚养权', '继承纠纷',
        '股权转让', '股东知情权', '公司决议', '损害公司利益',
        '建设工程', '房地产开发', '房屋买卖', '物业服务',
        '民间借贷', '金融借款', '保证合同',
    ]
    found = [kw for kw in dispute_keywords if kw in text]
    return found[:3]


def _extract_key_amounts(text):
    """Extract magnitude of key amounts (redacted, size only)."""
    amounts = []

    # Match: 人民币XXX万元
    mag_match = re.findall(r'人民币\s*[\d,]+\.?\d*\s*([万亿]?\s*元)', text)
    amounts.extend(mag_match[:3])

    # Match standalone amounts with magnitude
    mag_match2 = re.findall(r'[\d,]+\.?\d*\s*(万元|亿元)\b', text)
    amounts.extend(mag_match2[:3])

    # Deduplicate
    return list(set(amounts[:3]))


def _extract_court_level(text):
    """Detect court level from context (not from legal citations)."""
    # Look for "此致 XXXX人民法院" pattern (indicates actual court, not citation)
    court_match = re.search(r'此\s*致\s*\n?\s*([\u4e00-\u9fa5]+人民法院)', text)
    if court_match:
        court_name = court_match.group(1)
        if '最高' in court_name:
            return '最高人民法院'
        if '高级' in court_name:
            return '高级人民法院'
        if '中级' in court_name:
            return '中级人民法院'
        return '基层人民法院'

    # Fallback: look for actual court mentions in case headers (not citations)
    court_match2 = re.search(r'(?:受诉|管辖|受理)法院[：:]\s*([\u4e00-\u9fa5]+人民法院)', text)
    if court_match2:
        court_name = court_match2.group(1)
        if '中级' in court_name:
            return '中级人民法院'
        return '基层人民法院'

    return None


# ═══════════════════════════════════════════════════════════════
# DOCUMENT-TYPE-SPECIFIC ABSTRACTORS
# ═══════════════════════════════════════════════════════════════

def _abstract_contract(text, deid_info):
    """Generate abstract for a contract."""
    parties = {}
    for key, info in deid_info.get('parties', {}).items():
        # Determine if company or individual by name characteristics
        is_company = any(kw in info['name'] for kw in ['有限公司', '公司', '商行', '厂', '企业'])
        parties[key] = {
            'type': '公司/组织' if is_company else '自然人',
            'label': info['label'],
        }

    return {
        'doc_type': '合同/协议',
        'parties': parties,
        'clause_headings': _extract_headings(text),
        'clause_count_estimate': len(_extract_headings(text)),
        'significant_terms': [
            h for h in _extract_headings(text)
            if any(kw in h for kw in ['违约', '保密', '知识', '竞业', '终止', '解除', '管辖'])
        ],
        'blank_fields': ['公司名称', '注册资本', '签署日期'] if '_____' in text or '____' in text else [],
    }


def _abstract_complaint(text, deid_info):
    """Generate abstract for a complaint (起诉状)."""
    parties = {k: v['label'] for k, v in deid_info.get('parties', {}).items()}
    claims = _extract_headings(text)
    claim_indicators = re.findall(r'判令[^。；]{5,60}', text)[:5]

    return {
        'doc_type': '民事起诉状',
        'parties': parties,
        'claims': [c.strip()[:40] for c in claim_indicators] if claim_indicators else ['[详细诉讼请求已脱敏]'],
        'dispute_types': _extract_dispute_type(text),
        'legal_basis': _extract_legal_basis(text)[:5],
        'court': _extract_court_level(text),
    }


def _abstract_judgment(text, deid_info):
    """Generate abstract for a judgment (判决书)."""
    return {
        'doc_type': '民事判决书',
        'court': _extract_court_level(text),
        'dispute_focus': [
            h for h in _extract_headings(text)
            if any(kw in h for kw in ['争议', '焦点', '认定', '是否', '构成'])
        ][:5],
        'legal_basis': _extract_legal_basis(text)[:8],
        'dispute_types': _extract_dispute_type(text),
        'judgment_outcome': _extract_judgment_outcome(text),
    }


def _extract_judgment_outcome(text):
    """Extract judgment outcome from the ruling section (last ~1500 chars)."""
    tail = text[-1500:] if len(text) > 1500 else text
    # Check in reverse order of specificity
    if '部分支持' in tail:
        return '部分支持'
    if '驳回' in tail and ('诉讼请求' in tail or '上诉' in tail or '起诉' in tail):
        return '驳回'
    if '维持原判' in tail:
        return '维持原判'
    if '支持' in tail:
        return '支持原告诉讼请求'
    return '[裁判结果已脱敏]'


def _abstract_generic(text, deid_info):
    """Generate abstract for unrecognized document type."""
    parties = {k: v['label'] for k, v in deid_info.get('parties', {}).items()}
    return {
        'doc_type': '通用法律文书',
        'parties': parties if parties else None,
        'headings': _extract_headings(text)[:10],
        'legal_basis': _extract_legal_basis(text)[:5],
        'dispute_types': _extract_dispute_type(text),
    }


ABSTRACTORS = {
    '民事起诉状': _abstract_complaint,
    '民事答辩状': _abstract_generic,
    '民事判决书': _abstract_judgment,
    '民事裁定书': _abstract_judgment,
    '三方合作协议': _abstract_contract,
    '双方合同': _abstract_contract,
    '律师函': _abstract_generic,
    '法律意见书': _abstract_generic,
    '授权委托书': _abstract_generic,
    '仲裁申请书': _abstract_complaint,
    '上诉状': _abstract_complaint,
}


# ═══════════════════════════════════════════════════════════════
# MAIN API
# ═══════════════════════════════════════════════════════════════

def generate_abstract(text_or_deid_result, deid_info=None):
    """
    Generate a structured, de-identified abstract from legal document text.

    Args:
        text_or_deid_result: Either a raw text string OR the dict from deidentifier.deidentify()
        deid_info: Required if text is raw. The dict from deidentifier.deidentify()

    Returns:
        dict: Structured abstract safe for upload to AI API (~300-800 chars JSON)
    """
    # Accept either raw text or a deidentifier result dict
    if isinstance(text_or_deid_result, dict):
        deid_result = text_or_deid_result
        text = deid_result.get('deidentified_text', '')
        deid_info = deid_result
    else:
        text = text_or_deid_result

    if not text or not text.strip():
        return {'error': 'No text to abstract'}

    doc_type = detect_doc_type(text)
    abstractor = ABSTRACTORS.get(doc_type, _abstract_generic)

    abstract = abstractor(text, deid_info or {})
    abstract['detected_type'] = doc_type
    abstract['text_length'] = len(text)

    return abstract


def generate_abstract_from_file(filepath):
    """
    Full pipeline: read file → deidentify → abstract.
    Returns the abstract dict (safe to upload).
    """
    from deidentifier import deidentify_file
    deid = deidentify_file(filepath)
    return generate_abstract(deid)


def abstract_batch(filepaths, show_progress=False):
    """
    Batch process: deidentify + abstract for multiple files.

    Args:
        filepaths: List of file paths
        show_progress: Print progress to stderr (not to AI context)

    Returns:
        List of dicts, each: {
            'filename': str,
            'filepath': str,
            'doc_type': str,
            'parties': dict,
            'headings_or_claims': list,
            'legal_basis': list,
            'error': str or None,
            'original_chars': int,
        }
    """
    from deidentifier import deidentify_batch

    results = []
    deids = deidentify_batch(filepaths)

    for fp in filepaths:
        fname = os.path.basename(fp)
        deid = deids.get(fname, {})

        if 'error' in deid:
            results.append({
                'filename': fname, 'filepath': fp,
                'doc_type': '⚠ 无法识别', 'parties': {}, 'headings_or_claims': [],
                'legal_basis': [], 'error': deid['error'],
                'original_chars': 0
            })
            continue

        abstract = generate_abstract(deid)
        parties = {k: v.get('label', v) if isinstance(v, dict) else v
                   for k, v in abstract.get('parties', {}).items()}

        # Collect key items for display
        key_items = []
        if abstract.get('claims'):
            key_items = [c[:50] for c in abstract['claims'][:3]]
        elif abstract.get('clause_headings'):
            key_items = abstract['clause_headings'][:5]
        elif abstract.get('headings'):
            key_items = abstract['headings'][:5]

        results.append({
            'filename': fname, 'filepath': fp,
            'doc_type': abstract.get('detected_type', '未知'),
            'parties': parties,
            'headings_or_claims': key_items,
            'legal_basis': abstract.get('legal_basis', [])[:5],
            'error': None,
            'original_chars': deid.get('original_length', 0),
            'redactions': deid.get('redaction_count', 0),
        })

    return results


if __name__ == '__main__':
    # Test with contract sample
    contract_text = """
甲方：浙江清研生物科技有限公司
统一社会信用代码：91330108MA28ABCDEF
乙方：义乌市清研日用品商行
丙方：杭州深度科技有限公司

第一条 出资与持股
各方共同出资设立运营公司，甲方出资40%，乙方出资40%，丙方出资20%。

第二条 分工
甲方负责市场经营与运营管理，乙方负责外围资源对接，丙方负责全线技术支撑。

第三条 利润分配
40%固定分红 + 60%弹性分红（按贡献价值分配）。

第四条 违约责任
任何一方违反本协议约定，应支付违约金并赔偿全部损失。

第五条 争议解决
提交运营公司所在地人民法院诉讼解决。
"""
    from deidentifier import deidentify
    deid = deidentify(contract_text)
    abstract = generate_abstract(deid)

    # Manual serialization to avoid GBK terminal issues
    output = json.dumps(abstract, ensure_ascii=False, indent=2)
    print(output.encode('utf-8', errors='replace').decode('utf-8'))
    print(f'\nAbstract size: ~{len(output)} chars')
