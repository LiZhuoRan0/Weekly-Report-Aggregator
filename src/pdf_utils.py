"""PDF utilities: read text, count pages, validate, merge with bookmarks."""
import logging
from pathlib import Path
from typing import List, Optional, Tuple

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

logger = logging.getLogger("wra")


def is_valid_pdf(path: str) -> bool:
    """Return True if the file is a readable PDF, False if corrupt/unreadable."""
    try:
        reader = PdfReader(path)
        # Trigger parsing of at least one page if any
        _ = len(reader.pages)
        return True
    except (PdfReadError, OSError, ValueError, Exception) as e:
        logger.warning(f"Invalid/unreadable PDF: {path} ({e})")
        return False


def get_page_count(path: str) -> int:
    """Return number of pages in a PDF, or 0 if unreadable."""
    try:
        reader = PdfReader(path)
        return len(reader.pages)
    except Exception as e:
        logger.warning(f"Could not count pages in {path}: {e}")
        return 0


def extract_text_first_pages(path: str, max_pages: int = 2) -> str:
    """Extract text from up to the first `max_pages` pages of a PDF.

    Returns empty string on failure. Used for content-based name matching.
    """
    try:
        # Prefer pdfplumber for cleaner text, fall back to pypdf
        import pdfplumber
        text_chunks: List[str] = []
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages):
                if i >= max_pages:
                    break
                t = page.extract_text() or ""
                text_chunks.append(t)
        return "\n".join(text_chunks)
    except Exception as e:
        logger.debug(f"pdfplumber failed on {path}: {e}; falling back to pypdf")
        try:
            reader = PdfReader(path)
            chunks = []
            for i, page in enumerate(reader.pages):
                if i >= max_pages:
                    break
                chunks.append(page.extract_text() or "")
            return "\n".join(chunks)
        except Exception as e2:
            logger.warning(f"Failed to extract text from {path}: {e2}")
            return ""


def merge_pdfs_with_bookmarks(
    items: List[Tuple[str, str]],
    output_path: str,
) -> Optional[int]:
    """Merge PDFs and add a top-level bookmark for each.

    items: list of (bookmark_title, pdf_file_path) tuples in the desired order.
    output_path: where to write the merged PDF.
    Returns total page count, or None on total failure.
    """
    writer = PdfWriter()
    current_page = 0
    added_any = False

    for title, pdf_path in items:
        try:
            reader = PdfReader(pdf_path)
        except Exception as e:
            logger.error(f"Skipping unreadable PDF during merge: {pdf_path} ({e})")
            continue

        try:
            n_pages = len(reader.pages)
            if n_pages == 0:
                logger.warning(f"Skipping empty PDF: {pdf_path}")
                continue
            for page in reader.pages:
                writer.add_page(page)
            # Add bookmark pointing at the first page of this section
            writer.add_outline_item(title, current_page)
            current_page += n_pages
            added_any = True
        except Exception as e:
            logger.error(f"Error appending {pdf_path} to merge: {e}")
            continue

    if not added_any:
        logger.error("No PDFs were successfully merged.")
        return None

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(output_path, "wb") as f:
            writer.write(f)
    except Exception as e:
        logger.error(f"Failed to write merged PDF to {output_path}: {e}")
        return None

    return current_page


def split_merged_pdf_by_size(
    items: List[Tuple[str, str]],
    output_dir: str,
    base_name: str,
    max_size_bytes: int,
) -> List[str]:
    """Bin-pack student PDFs into chunks each <= max_size_bytes, merging each chunk.

    items: ordered list of (bookmark_title, pdf_file_path).
    Returns list of generated chunk paths. If all items fit in one chunk, returns
    a single-element list.

    Each chunk preserves student order. Bookmarks are scoped within each chunk.
    Filenames: {base_name}.pdf for single, or {base_name}_part{i}of{n}.pdf for multi.
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # First, get the size of each item's source file
    sized: List[Tuple[str, str, int]] = []
    for title, path in items:
        try:
            sz = Path(path).stat().st_size
        except OSError as e:
            logger.warning(f"Cannot stat {path}: {e}; skipping in size estimation")
            continue
        sized.append((title, path, sz))

    if not sized:
        logger.error("No PDFs available to split/merge.")
        return []

    # Bin-pack greedily into chunks. Note: merged size is roughly sum of inputs;
    # we leave ~5% headroom for PDF overhead.
    headroom = int(max_size_bytes * 0.95)

    chunks: List[List[Tuple[str, str]]] = []
    current_chunk: List[Tuple[str, str]] = []
    current_size = 0
    for title, path, sz in sized:
        if sz > max_size_bytes:
            # Single file already exceeds the limit. We still include it alone.
            logger.warning(
                f"PDF '{path}' is {sz} bytes which exceeds the per-email limit. "
                f"It will be sent alone in its own chunk."
            )
            if current_chunk:
                chunks.append(current_chunk)
                current_chunk = []
                current_size = 0
            chunks.append([(title, path)])
            continue
        if current_size + sz > headroom and current_chunk:
            chunks.append(current_chunk)
            current_chunk = []
            current_size = 0
        current_chunk.append((title, path))
        current_size += sz
    if current_chunk:
        chunks.append(current_chunk)

    n = len(chunks)
    output_paths: List[str] = []
    for i, chunk in enumerate(chunks, start=1):
        if n == 1:
            out_path = out_dir / f"{base_name}.pdf"
        else:
            out_path = out_dir / f"{base_name}_part{i}of{n}.pdf"
        result = merge_pdfs_with_bookmarks(chunk, str(out_path))
        if result is None:
            logger.error(f"Failed to build chunk {i}/{n}")
            continue
        output_paths.append(str(out_path))
        logger.info(f"Wrote chunk {i}/{n}: {out_path} ({Path(out_path).stat().st_size / 1024 / 1024:.2f} MB)")

    return output_paths
