"""
Reusable contract image processing utilities for bid document generation.

Core functions:
  - make_contract_collage(): Extract front+signature pages as side-by-side PNG
  - render_full_page(): Render a single PDF page at target DPI
  - batch_collage(): Process multiple contracts in one call
  - find_signature_page(): OCR-based detection (placeholder for future)
"""

import os, sys, fitz
from PIL import Image


def render_page_to_png(pdf_path, page_num, dpi=200):
    """Render a single PDF page to a PIL Image."""
    doc = fitz.open(pdf_path)
    if page_num < 0:
        page_num = len(doc) + page_num
    page = doc[page_num]
    pix = page.get_pixmap(dpi=dpi)
    # Save to temp then reload (PyMuPDF pixmap -> PIL)
    import io
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    doc.close()
    return img


def make_contract_collage(pdf_path, output_path=None, dpi=180, gap=10):
    """
    Render front page + last page (signature page) of a contract PDF
    as a side-by-side collage image.

    Args:
        pdf_path: Path to contract PDF
        output_path: Where to save the collage PNG (default: same dir, _collage suffix)
        dpi: Render resolution
        gap: Gap in pixels between the two pages

    Returns:
        Path to the saved collage image, or None on failure.
    """
    try:
        front_img = render_page_to_png(pdf_path, 0, dpi)
        last_img = render_page_to_png(pdf_path, -1, dpi)

        # Resize to matching height
        h = max(front_img.height, last_img.height)
        w0 = int(front_img.width * h / front_img.height)
        w1 = int(last_img.width * h / last_img.height)
        front_img = front_img.resize((w0, h), Image.LANCZOS)
        last_img = last_img.resize((w1, h), Image.LANCZOS)

        # Side-by-side collage
        total_w = w0 + w1 + gap
        collage = Image.new('RGB', (total_w, h), (255, 255, 255))
        collage.paste(front_img, (0, 0))
        collage.paste(last_img, (w0 + gap, 0))

        if output_path is None:
            base = os.path.splitext(pdf_path)[0]
            output_path = f"{base}_collage.png"

        collage.save(output_path)
        return output_path

    except Exception as e:
        print(f"[contract_utils] Collage failed for {pdf_path}: {e}", file=sys.stderr)
        return None


def render_full_page(pdf_path, output_path, page_num=0, dpi=200):
    """Render a single full page of a PDF and save as PNG."""
    try:
        img = render_page_to_png(pdf_path, page_num, dpi)
        img.save(output_path)
        return output_path
    except Exception as e:
        print(f"[contract_utils] Page render failed: {e}", file=sys.stderr)
        return None


def batch_collage(contracts, output_dir, dpi=180, gap=10):
    """
    Process a batch of contracts, generating collages for each.

    Args:
        contracts: List of (pdf_path, label) tuples
        output_dir: Directory for output PNGs
        dpi: Render resolution

    Returns:
        Dict mapping label -> output path (or None on failure)
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {}
    for pdf_path, label in contracts:
        out_path = os.path.join(output_dir, f"collage_{label}.png")
        result = make_contract_collage(pdf_path, out_path, dpi, gap)
        results[label] = result
        status = "OK" if result else "FAIL"
        print(f"  [{status}] {label} -> {out_path}", file=sys.stderr)
    return results


def render_brochure_pages(pdf_path, page_numbers, output_dir, prefix='brochure', dpi=200):
    """
    Render and save specific pages from a promotional brochure.

    Args:
        pdf_path: Path to brochure PDF
        page_numbers: List of 0-indexed page numbers
        output_dir: Output directory
        prefix: Filename prefix
        dpi: Render resolution

    Returns:
        List of output file paths.
    """
    os.makedirs(output_dir, exist_ok=True)
    results = []
    doc = fitz.open(pdf_path)
    for pn in page_numbers:
        if pn < len(doc):
            out_path = os.path.join(output_dir, f'{prefix}_P{pn+1}.png')
            page = doc[pn]
            pix = page.get_pixmap(dpi=dpi)
            pix.save(out_path)
            results.append(out_path)
            print(f"  Brochure P{pn+1} -> {out_path}", file=sys.stderr)
    doc.close()
    return results


# ============================================================
# DOCX SIZE ESTIMATION & CHAPTER MERGE
# ============================================================

def estimate_docx_size(image_paths, text_paras=0, dpi=200):
    """
    Estimate final DOCX size in MB.
    
    Rough benchmarks:
      - Collage @200dpi: ~2.5 MB each
      - Single page @200dpi: ~1.0 MB each
      - Invoice JPG: ~0.1 MB each
      - Full-text page (no images): ~0.05 MB each
    
    Args:
        image_paths: List of image file paths (or PIL Images)
        text_paras: Approximate number of text-only paragraphs
        dpi: Render DPI (for benchmarking)
    
    Returns:
        Estimated size in MB (float)
    """
    size_mb = 0
    dpi_factor = dpi / 200.0
    
    for p in image_paths:
        if isinstance(p, str) and os.path.exists(p):
            fsize = os.path.getsize(p) / (1024 * 1024)
            # Collages are large, single pages smaller, invoices tiny
            if 'collage' in os.path.basename(p).lower():
                size_mb += 2.5 * dpi_factor
            elif '发票' in p or 'invoice' in os.path.basename(p).lower():
                size_mb += 0.1
            else:
                size_mb += fsize * 0.8  # PNG -> DOCX compression
        else:
            size_mb += 1.0  # Unknown image, assume single page
    
    # Text overhead
    size_mb += text_paras * 0.05 / 200  # ~200 paras = 0.05 MB
    
    return round(size_mb, 1)


def estimate_chapter_sizes(chapters):
    """
    Given a dict of {chapter_name: [image_paths]}, estimate per-chapter sizes.
    
    Args:
        chapters: Dict mapping chapter label to list of image paths.
    
    Returns:
        Dict mapping chapter label to estimated size in MB.
    """
    return {name: estimate_docx_size(imgs) for name, imgs in chapters.items()}


def merge_chapters(chapter_paths, output_path, keep_cover_only_first=True):
    """
    Merge multiple chapter DOCX files into one.
    
    Each chapter DOCX has a cover page as its first content. When merging,
    only the first chapter's cover page is kept; subsequent chapter cover
    pages are skipped.
    
    Args:
        chapter_paths: List of DOCX file paths in merge order.
        output_path: Where to save the merged DOCX.
        keep_cover_only_first: If True, skip cover pages for chapters 2+.
    
    Returns:
        output_path on success, None on failure.
    """
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.shared import Cm
    import copy

    try:
        merged = None

        for idx, ch_path in enumerate(chapter_paths):
            if not os.path.exists(ch_path):
                print(f"  SKIP missing: {ch_path}", file=sys.stderr)
                continue

            ch_doc = Document(ch_path)

            if merged is None:
                # First chapter: copy whole document as-is
                merged = ch_doc
                # Remove the trailing blank page break if present
                merged.element.body.remove(merged.element.body[-1])
            else:
                # Subsequent chapters: skip cover page, copy rest
                elements = list(ch_doc.element.body)
                start_idx = 0

                if keep_cover_only_first and idx > 0:
                    # Skip elements until first page break (cover page boundaries)
                    for ei, el in enumerate(elements):
                        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                        if tag == 'p':
                            # Check for page break in this paragraph
                            pPr = el.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                            if pPr is not None:
                                sectPr = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
                                if sectPr is not None:
                                    start_idx = ei + 1
                                    break
                            # Also check for lastRenderedPageBreak or actual page break
                            rPr = el.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastRenderedPageBreak')
                            pb = el.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
                            if rPr is not None or pb is not None:
                                # Simple heuristic: skip first page if it has sections + second page
                                pass

                        if tag == 'tbl':
                            start_idx = ei
                            break

                # Copy elements from start_idx to merged document
                for el in elements[start_idx:]:
                    merged.element.body.append(copy.deepcopy(el))

        if merged is not None:
            merged.save(output_path)
            print(f"Merged {len(chapter_paths)} chapters -> {output_path}", file=sys.stderr)
            return output_path
        else:
            print("merge_chapters: no valid input files", file=sys.stderr)
            return None

    except Exception as e:
        print(f"merge_chapters failed: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return None


def suggest_split_plan(chapters, threshold_mb=30):
    """
    Suggest whether to split output based on total estimated size.
    
    Args:
        chapters: Dict mapping chapter name to estimated size in MB.
        threshold_mb: Suggest split if total exceeds this.
    
    Returns:
        Dict with keys: 'split' (bool), 'total_mb' (float), 'chapter_count' (int),
        'plan' (list of chapter names in order)
    """
    total = sum(chapters.values())
    return {
        'split': total > threshold_mb,
        'total_mb': total,
        'chapter_count': len(chapters),
        'plan': list(chapters.keys()),
        'per_chapter_mb': {k: round(v, 1) for k, v in chapters.items()},
    }
