"""
Directory scanner for legal document folders.
Every file undergoes the full deidentifier→abstractor pipeline locally.
Only the file index (labels, not content) enters AI context.

Usage:
    from directory_scanner import scan_directory, format_index_for_ai
    index = scan_directory(r'C:\案卷\')
    print(format_index_for_ai(index))  # safe for AI context
"""

import os


SUPPORTED_EXTENSIONS = {'.docx', '.pdf', '.txt'}
IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp'}
SPREADSHEET_EXTENSIONS = {'.xlsx', '.xlsm', '.xls', '.csv'}
UNSUPPORTED_BUT_NOTABLE = {'.doc', '.pptx', '.ppt', '.zip', '.rar', '.7z'}


def _extract_label(file_path):
    """Run full deidentifier→abstract pipeline on a single file and return a
    compact human-readable label. ALL LOCAL — no content enters AI context."""
    import sys
    dir_path = os.path.dirname(os.path.abspath(__file__))
    if dir_path not in sys.path:
        sys.path.insert(0, dir_path)
    from deidentifier import deidentify_file
    from abstractor import generate_abstract

    try:
        deid = deidentify_file(file_path)
        abstract = generate_abstract(deid)

        doc_type = abstract.get('detected_type', '未知')
        parties = abstract.get('parties', {})
        party_tags = []
        for v in parties.values():
            # v can be a string label or a dict with 'label' key
            label = v if isinstance(v, str) else v.get('label', str(v))
            if label and label not in party_tags:
                party_tags.append(label)

        parts = [doc_type]
        if party_tags:
            parts.append(f"[{'/'.join(party_tags[:3])}]")

        if abstract.get('dispute_types'):
            parts.append(abstract['dispute_types'][0])
        elif abstract.get('claims'):
            claim = abstract['claims'][0][:30].strip()
            if claim:
                parts.append(claim)
        elif abstract.get('clause_headings'):
            parts.append(f'{len(abstract["clause_headings"])}条款')

        if abstract.get('legal_basis'):
            parts.append(abstract['legal_basis'][0][:15])

        return ' · '.join(parts)

    except Exception as e:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in IMAGE_EXTENSIONS:
            return f'🖼 图片 [{ext}] - 需OCR处理'
        if ext in SPREADSHEET_EXTENSIONS:
            return f'📊 表格 [{ext}] - 可能为数据或证据'
        return f'⚠ {os.path.basename(file_path)} [{ext}] - 无法识别'


def scan_directory(path):
    """
    Scan a directory. Every supported file runs the full pipeline locally
    to extract document type and party labels.

    ALL LOCAL — only labels enter AI context. Original file content never
    leaves the machine.

    Args:
        path: Directory path

    Returns:
        dict with 'path', 'file_count', 'total_size_mb', 'files'
        Each file entry: {filename, filepath, ext, size_mb, label, supported}
    """
    if not os.path.isdir(path):
        return {'error': f'Not a directory: {path}', 'path': path}

    files = []
    total_size = 0

    try:
        entries = list(os.scandir(path))
    except (PermissionError, FileNotFoundError, OSError) as e:
        return {'error': f'Cannot read directory: {e}', 'path': path}

    for entry in entries:
        if not entry.is_file() or entry.name.startswith('~$'):
            continue
        fname = entry.name
        fpath = entry.path
        ext = os.path.splitext(fname)[1].lower()
        try:
            size = entry.stat().st_size / (1024 * 1024)
        except (PermissionError, FileNotFoundError, OSError):
            size = 0
        total_size += size
        is_supported = ext in SUPPORTED_EXTENSIONS

        files.append({
            'filename': fname,
            'filepath': fpath,
            'ext': ext,
            'size_mb': round(size, 2),
            'supported': is_supported,
                'is_image': ext in IMAGE_EXTENSIONS,
                'is_spreadsheet': ext in SPREADSHEET_EXTENSIONS,
                'label': None,
            })

    # Run pipeline on EVERY supported file — no filename shortcuts
    for f in files:
        if f['supported']:
            f['label'] = _extract_label(f['filepath'])
        elif f['is_image']:
            f['label'] = f'🖼 {f["filename"]} - 需OCR处理'
        elif f['is_spreadsheet']:
            f['label'] = f'📊 {f["filename"]} - 表格文件'
        elif f['ext'] in UNSUPPORTED_BUT_NOTABLE:
            f['label'] = f'📎 {f["filename"]} - 需转换格式'
        else:
            f['label'] = f'📄 {f["filename"]}'

    # Sort: supported first, then images/spreadsheets, then unsupported
    files.sort(key=lambda f: (not f['supported'], not f['is_image'], f['filename']))

    return {
        'path': path,
        'file_count': len(files),
        'total_size_mb': round(total_size, 2),
        'files': files,
    }


def format_index_for_ai(index):
    """
    Convert scan_directory() result into a Markdown summary safe for AI context.
    Each file line is ~30-80 bytes of document type label — zero content exposure.

    Args:
        index: Result from scan_directory()

    Returns:
        Formatted Markdown string
    """
    if not isinstance(index, dict):
        return f'**Error:** invalid index value: {type(index).__name__}'
    if 'error' in index:
        return f'**Error:** {index["error"]}'

    path = index.get('path', '?')
    file_count = index.get('file_count', 0)
    total_mb = index.get('total_size_mb', 0)
    files = index.get('files', [])

    lines = [
        f'📁 **{os.path.basename(path)}** '
        f'({file_count} 个文件, {total_mb} MB)',
        '',
    ]

    for i, f in enumerate(files, 1):
        size_str = f'{f["size_mb"]:.1f}MB' if f['size_mb'] >= 0.1 else '<0.1MB'
        label = f['label'] or f['filename']
        status = ''
        if not f['supported'] and not f['is_image'] and not f['is_spreadsheet']:
            status = ' ⚠需转换'
        elif f['is_image']:
            status = ' 🖼需OCR'
        lines.append(f'  {i}. {label}{status} [{size_str}]')

    return '\n'.join(lines)


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("Usage: python directory_scanner.py <directory_path>")
        sys.exit(1)

    idx = scan_directory(sys.argv[1])
    print(format_index_for_ai(idx))
