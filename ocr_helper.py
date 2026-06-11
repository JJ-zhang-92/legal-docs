"""
OCR helper for scanned legal documents.
Supports scanned PDFs, images (PNG/JPG/TIFF), and mixed documents.

Prerequisites:
    pip install pytesseract Pillow pdf2image
    Download and install Tesseract-OCR with Chinese language pack:
    https://github.com/UB-Mannheim/tesseract/wiki (Windows)
    macOS: brew install tesseract tesseract-lang
    Linux: apt install tesseract-ocr tesseract-ocr-chi-sim
"""

import os
import sys
import tempfile
from pathlib import Path


def _deidentify_if_available(text):
    """Auto-deidentify OCR output before returning. When called from
    within the legal-docs skill, deidentifier is importable and all
    OCR text is sanitized before reaching AI context. Standalone use
    outside the skill skips deidentification."""
    try:
        skill_dir = os.path.dirname(os.path.abspath(__file__))
        if skill_dir not in sys.path:
            sys.path.insert(0, skill_dir)
        from deidentifier import deidentify
        return deidentify(text)['deidentified_text']
    except ImportError:
        return text


def ensure_tesseract():
    """Check if Tesseract is installed and configured. Return (ok, message)."""
    import subprocess
    try:
        result = subprocess.run(
            ['tesseract', '--version'],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            # Check Chinese language pack
            lang_check = subprocess.run(
                ['tesseract', '--list-langs'],
                capture_output=True, text=True, timeout=10
            )
            has_chi_sim = 'chi_sim' in lang_check.stdout
            return True, f"Tesseract OK (chi_sim={'YES' if has_chi_sim else 'NO - install Chinese lang pack!'})"
        return False, f"Tesseract error: {result.stderr}"
    except FileNotFoundError:
        return False, (
            "Tesseract not installed.\n"
            "Windows: download from https://github.com/UB-Mannheim/tesseract/wiki\n"
            "macOS: brew install tesseract tesseract-lang\n"
            "Linux: sudo apt install tesseract-ocr tesseract-ocr-chi-sim\n"
            "After install, set TESSERACT_PATH in environment or ensure it's in PATH."
        )


def ocr_image(image_path, lang='chi_sim+eng', config=''):
    """
    Extract text from a single image file.

    Args:
        image_path: Path to image file (PNG, JPG, TIFF, BMP)
        lang: Tesseract language string ('chi_sim' for simplified Chinese,
              'chi_sim+eng' for Chinese+English, 'chi_tra' for traditional)
        config: Additional Tesseract config (e.g. '--psm 6' for uniform text block)

    Returns:
        Extracted text string.
    """
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return _install_instructions('pytesseract')

    ok, msg = ensure_tesseract()
    if not ok:
        return f"OCR NOT AVAILABLE: {msg}"

    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang=lang, config=config)
    return _deidentify_if_available(text)


def ocr_pdf(pdf_path, lang='chi_sim+eng', dpi=300, page_range=None):
    """
    Extract text from a scanned PDF using OCR.

    Args:
        pdf_path: Path to PDF file
        lang: Tesseract language
        dpi: Resolution for PDF→image conversion (higher = better OCR, slower)
        page_range: Tuple (start, end) for specific pages, or None for all

    Returns:
        Dict with keys: 'pages' (list of {page_num, text}), 'full_text' (combined)
    """
    try:
        from pdf2image import convert_from_path
    except ImportError:
        return {'error': _install_instructions('pdf2image')}

    ok, msg = ensure_tesseract()
    if not ok:
        return {'error': f"OCR NOT AVAILABLE: {msg}"}

    import pytesseract

    # Convert PDF to images
    kwargs = {'dpi': dpi}
    if page_range:
        kwargs['first_page'] = page_range[0]
        kwargs['last_page'] = page_range[1]

    images = convert_from_path(pdf_path, **kwargs)

    pages = []
    all_text = []

    for i, img in enumerate(images):
        page_num = (page_range[0] + i) if page_range else (i + 1)
        text = pytesseract.image_to_string(img, lang=lang)
        text = _deidentify_if_available(text)
        pages.append({'page_num': page_num, 'text': text})
        all_text.append(f"--- 第 {page_num} 页 ---\n{text}")

    return {
        'pages': pages,
        'full_text': _deidentify_if_available('\n'.join(all_text)),
        'page_count': len(pages),
    }


def smart_ocr(file_path, lang='chi_sim+eng'):
    """
    Auto-detect file type and apply appropriate OCR.

    For PDFs: checks if document has extractable text first (digital PDF),
    only OCRs scanned pages without text.

    Args:
        file_path: Path to PDF or image file
        lang: OCR language

    Returns:
        Dict with extracted text and metadata.
    """
    ext = Path(file_path).suffix.lower()

    if ext == '.pdf':
        # First try extracting embedded text
        try:
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                text_pages = []
                empty_pages = []
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text()
                    if t and t.strip():
                        text_pages.append({'page_num': i+1, 'text': t.strip()})
                    else:
                        empty_pages.append(i+1)

            if empty_pages and len(empty_pages) == len(text_pages) + len(empty_pages):
                # All pages are scanned — OCR everything
                return ocr_pdf(file_path, lang=lang)
            elif empty_pages:
                # Mix — OCR only empty pages
                ocr_result = ocr_pdf(file_path, lang=lang,
                                     page_range=(min(empty_pages), max(empty_pages)))
                # Merge
                deid_text_pages = []
                for p in text_pages:
                    deid_text_pages.append({'page_num': p['page_num'],
                                            'text': _deidentify_if_available(p['text'])})
                all_pages = {p['page_num']: p['text'] for p in deid_text_pages}
                for p in ocr_result.get('pages', []):
                    all_pages[p['page_num']] = p['text']

                merged = []
                for pg in sorted(all_pages.keys()):
                    merged.append(f"--- 第 {pg} 页 ---\n{all_pages[pg]}")

                return {
                    'pages': [{'page_num': pg, 'text': all_pages[pg]} for pg in sorted(all_pages.keys())],
                    'full_text': '\n'.join(merged),
                    'page_count': len(all_pages),
                    'ocr_pages': empty_pages,
                }
            else:
                # All text extracted — no OCR needed
                deid_pages = []
                all_lines = []
                for i, p in enumerate(text_pages):
                    deid_text = _deidentify_if_available(p['text'])
                    deid_pages.append({'page_num': p['page_num'], 'text': deid_text})
                    all_lines.append(f"--- 第 {i+1} 页 ---\n{deid_text}")
                return {
                    'pages': deid_pages,
                    'full_text': '\n'.join(all_lines),
                    'page_count': len(deid_pages),
                    'ocr_pages': [],
                }
        except ImportError:
            # pdfplumber not available, fallback to full OCR
            return ocr_pdf(file_path, lang=lang)

    elif ext in ('.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.gif'):
        text = ocr_image(file_path, lang=lang)
        return {
            'pages': [{'page_num': 1, 'text': text}],
            'full_text': text,
            'page_count': 1,
        }
    else:
        return {'error': f"Unsupported file type: {ext}"}


def extract_with_preprocessing(image_path, output_dir=None):
    """
    Enhanced OCR with image preprocessing for low-quality scans.

    Applies: grayscale → denoise → threshold (binarize) → deskew → OCR

    Args:
        image_path: Path to the scanned image/PDF
        output_dir: Optional directory to save preprocessed intermediate images

    Returns:
        Extracted text string.
    """
    try:
        from PIL import Image, ImageFilter, ImageEnhance
        import numpy as np
    except ImportError:
        return _install_instructions('Pillow numpy')

    ok, msg = ensure_tesseract()
    if not ok:
        return f"OCR NOT AVAILABLE: {msg}"

    import pytesseract

    img = Image.open(image_path)

    # 1. Convert to grayscale
    img = img.convert('L')

    # 2. Enhance contrast
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)

    # 3. Denoise
    img = img.filter(ImageFilter.MedianFilter(size=3))

    # 4. Binarize (threshold)
    threshold = 128
    img = img.point(lambda p: 255 if p > threshold else 0)

    # 5. Save intermediate if requested
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        preprocessed_path = os.path.join(output_dir, 'preprocessed.png')
        img.save(preprocessed_path)
        print(f"预处理图片已保存: {preprocessed_path}")

    # OCR with layout analysis
    text = pytesseract.image_to_string(
        img, lang='chi_sim+eng',
        config='--psm 6 -c preserve_interword_spaces=1'
    )

    return _deidentify_if_available(text)


def batch_ocr_directory(directory, output_file=None, recursive=True):
    """
    OCR all scanned documents in a directory.

    Args:
        directory: Path to search
        output_file: Optional output text file for combined results
        recursive: Whether to search subdirectories

    Returns:
        Dict mapping filename → extracted text
    """
    results = {}

    image_extensions = {'.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.pdf'}

    if recursive:
        for root, dirs, files in os.walk(directory):
            for f in files:
                ext = Path(f).suffix.lower()
                if ext in image_extensions:
                    filepath = os.path.join(root, f)
                    print(f"处理: {filepath}", file=sys.stderr)
                    result = smart_ocr(filepath)
                    results[filepath] = result.get('full_text', result.get('error', ''))
    else:
        for f in os.listdir(directory):
            ext = Path(f).suffix.lower()
            if ext in image_extensions:
                filepath = os.path.join(directory, f)
                print(f"处理: {filepath}", file=sys.stderr)
                result = smart_ocr(filepath)
                results[filepath] = result.get('full_text', result.get('error', ''))

    if output_file:
        with open(output_file, 'w', encoding='utf-8') as out:
            for filepath, text in results.items():
                out.write(f"\n{'='*60}\n")
                out.write(f"文件: {filepath}\n")
                out.write(f"{'='*60}\n\n")
                out.write(text)
                out.write('\n')
        print(f"批量OCR结果已保存至: {output_file}")

    return results


def _install_instructions(package):
    return (
        f"Missing package: {package}\n"
        f"Run: pip install {package}\n\n"
        f"For full OCR setup:\n"
        f"  1. Install Tesseract-OCR:\n"
        f"     Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
        f"     macOS: brew install tesseract tesseract-lang\n"
        f"     Linux: sudo apt install tesseract-ocr tesseract-ocr-chi-sim\n"
        f"  2. pip install pytesseract Pillow pdf2image\n"
        f"  3. If on Windows, set TESSDATA_PREFIX or add Tesseract to PATH"
    )


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python ocr_helper.py <文件路径> [语言]")
        print("示例: python ocr_helper.py 扫描件.pdf")
        print("      python ocr_helper.py 合同扫描.pdf chi_sim+eng")
        print("      python ocr_helper.py --batch 合同文件夹/")
        print("      python ocr_helper.py --check")
        sys.exit(1)

    if sys.argv[1] == '--check':
        ok, msg = ensure_tesseract()
        print(f"Tesseract: {'OK' if ok else 'NOT OK'}")
        print(msg)
    elif sys.argv[1] == '--batch':
        batch_dir = sys.argv[2] if len(sys.argv) > 2 else '.'
        batch_ocr_directory(batch_dir, output_file='ocr_results.txt')
    else:
        file_path = sys.argv[1]
        lang = sys.argv[2] if len(sys.argv) > 2 else 'chi_sim+eng'
        result = smart_ocr(file_path, lang=lang)
        if 'error' in result:
            print(f"错误: {result['error']}")
        else:
            print(result.get('full_text', ''))
