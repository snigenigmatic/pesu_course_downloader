"""
PDF Duplicate Detection and Cleanup
====================================
Drop-in function for the PESU Academy Downloader.

Dependencies (add to your pip installs):
    pip install pymupdf   # fitz — for rendering PDF pages to images

How it works:
  1. For each PDF in a resource folder, render a sample of "middle" slides
     (skipping first 2 and last 2 — intro/thank-you slides) to grayscale images.
  2. Compute a perceptual hash (pHash) for each sampled page.
  3. Two PDFs are considered duplicates if their average pHash distance is below
     a threshold (default: 8 bits out of 64 — very lenient for slide decks).
  4. Within each duplicate group, keep the first file (lowest numbered), delete rest.
  5. After deletion, renumber remaining files sequentially (only the leading number
     in the filename, e.g. "3.SomeTopic.pdf" → "2.SomeTopic.pdf").

Insert point in main():
  After the convert_office_to_pdf() call (or after downloads if no conversion),
  before merge_pdfs_by_type(), add:

      deduplicate_pdfs_in_folder(base_dir, selected_resources)
"""

import math
import re
from pathlib import Path
from typing import Dict, List, Optional

from colorama import Fore, Style

# ── pHash helpers (no extra deps beyond pymupdf) ─────────────────────────────

def _render_page_gray(page, size: int = 32):
    """Render a PDF page to a grayscale size×size pixel list using PyMuPDF."""
    import fitz  # PyMuPDF
    mat = fitz.Matrix(size / page.rect.width, size / page.rect.height)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
    return list(pix.samples)  # flat list of 0-255 ints, length = size*size


def _phash(pixels: List[int], size: int = 32) -> int:
    """
    Compute a 64-bit perceptual hash from a flat grayscale pixel list.
    Uses 8×8 DCT-based pHash logic on the size×size input.
    Returns an integer (the hash bits packed).
    """
    N = size
    # Compute 8×8 DCT of the image
    dct_size = 8
    pixels_2d = [[pixels[r * N + c] for c in range(N)] for r in range(N)]

    dct = [[0.0] * dct_size for _ in range(dct_size)]
    for u in range(dct_size):
        for v in range(dct_size):
            s = 0.0
            cu = math.sqrt(0.5) if u == 0 else 1.0
            cv = math.sqrt(0.5) if v == 0 else 1.0
            for x in range(N):
                for y in range(N):
                    s += (pixels_2d[x][y] *
                          math.cos((2 * x + 1) * u * math.pi / (2 * N)) *
                          math.cos((2 * y + 1) * v * math.pi / (2 * N)))
            dct[u][v] = cu * cv * s

    # Flatten and compute median (exclude [0][0] DC component)
    flat = [dct[u][v] for u in range(dct_size) for v in range(dct_size)][1:]
    median = sorted(flat)[len(flat) // 2]

    # Build hash: 1 if above median, 0 otherwise
    bits = [(1 if dct[u][v] > median else 0)
            for u in range(dct_size) for v in range(dct_size)]

    h = 0
    for b in bits:
        h = (h << 1) | b
    return h


def _hamming(a: int, b: int) -> int:
    """Bit-count of XOR — number of differing bits."""
    return bin(a ^ b).count("1")


def _pdf_fingerprint(pdf_path: Path,
                     sample_count: int = 4,
                     skip_edges: int = 2) -> Optional[List[int]]:
    """
    Return a list of pHashes for `sample_count` interior pages of a PDF.
    Skips the first and last `skip_edges` pages.
    Returns None if the PDF can't be opened or has too few pages.
    """
    try:
        import fitz
        with fitz.open(str(pdf_path)) as doc:
            total = len(doc)

            # Need enough interior pages
            interior_start = skip_edges
            interior_end   = total - skip_edges  # exclusive
            interior_count = interior_end - interior_start

            if interior_count <= 0:
                # PDF is tiny — just hash all pages
                page_indices = list(range(total))
            else:
                # Evenly-spaced sample across the interior
                if interior_count <= sample_count:
                    page_indices = list(range(interior_start, interior_end))
                else:
                    step = interior_count / sample_count
                    page_indices = [
                        interior_start + int(i * step)
                        for i in range(sample_count)
                    ]

            hashes = []
            for idx in page_indices:
                page   = doc[idx]
                pixels = _render_page_gray(page, size=32)
                hashes.append(_phash(pixels))

        return hashes if hashes else None

    except Exception as e:
        print(f"      {Fore.YELLOW}⚠ Could not fingerprint {pdf_path.name}: {e}{Style.RESET_ALL}")
        return None


def _are_duplicates(hashes_a: List[int],
                    hashes_b: List[int],
                    threshold: int = 8) -> bool:
    """
    Two PDFs are duplicates if the average Hamming distance between their
    corresponding page hashes is ≤ threshold bits (out of 64).
    """
    if len(hashes_a) != len(hashes_b):
        # Different sample lengths — compare the minimum overlap
        pairs = list(zip(hashes_a, hashes_b))
        if not pairs:
            return False
    else:
        pairs = list(zip(hashes_a, hashes_b))

    avg_dist = sum(_hamming(a, b) for a, b in pairs) / len(pairs)
    return avg_dist <= threshold


def _natural_sort_key(path: Path) -> list:
    """Numeric-aware sort key (same as in your main script)."""
    parts = re.split(r'(\d+)', path.name)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def _leading_number(filename: str) -> Optional[int]:
    """Extract the leading number from a filename like '3.SomeTopic.pdf'."""
    m = re.match(r'^(\d+)[\._]', filename)
    return int(m.group(1)) if m else None


def _renumber_files(pdf_files: List[Path]) -> List[Path]:
    """
    Renumber the leading integer in filenames to be sequential (1, 2, 3, …).
    Files that have no leading number are left untouched.
    Uses two-phase rename (via temp names) to avoid collisions on Windows.
    Returns the new paths.
    """
    numbered = [(p, _leading_number(p.name)) for p in pdf_files]
    # Sort by current number (None last)
    numbered.sort(key=lambda x: (x[1] is None, x[1] or 0))

    # Phase 1: rename all numbered files to unique temp names to avoid collisions
    # e.g. "3.Topic.pdf" → "__tmp_3.Topic.pdf"
    temp_numbered = []
    for path, old_num in numbered:
        if old_num is None:
            temp_numbered.append((path, None))
            continue
        tmp_path = path.parent / f"__tmp_{path.name}"
        path.rename(tmp_path)
        temp_numbered.append((tmp_path, old_num))

    # Phase 2: rename temp files to final sequential names
    new_paths = []
    counter   = 1

    for tmp_path, old_num in temp_numbered:
        if old_num is None:
            new_paths.append(tmp_path)
            continue

        new_name = re.sub(r'^\d+', str(counter), tmp_path.name.replace("__tmp_", "", 1), count=1)
        new_path = tmp_path.parent / new_name
        tmp_path.rename(new_path)

        new_paths.append(new_path)
        counter += 1

    return new_paths


# ── Main public function ───────────────────────────────────────────────────────

RESOURCE_TYPES = {          # keep in sync with your main script
    "2": "Slides",
    "3": "Notes",
    "4": "QA",
    "5": "Assignments",
    "6": "QB",
    "7": "MCQs",
    "8": "References",
}


def deduplicate_pdfs_in_folder(base_dir: Path,
                                resource_type_ids: List[str],
                                sample_count: int = 4,
                                skip_edges: int = 2,
                                hash_threshold: int = 8,
                                auto_delete: bool = False) -> None:
    """
    Scan every Unit_*/ResourceType/ folder, detect duplicate PDFs by perceptual
    hashing, prompt the user for action, then renumber surviving files.

    Parameters
    ----------
    base_dir            : root download folder (e.g. Path("downloads/CourseName"))
    resource_type_ids   : list of resource IDs selected by user (e.g. ["2","6"])
    sample_count        : interior pages to sample per PDF (default 4)
    skip_edges          : pages to skip at start/end (intro + thank-you slides)
    hash_threshold      : max average Hamming distance to call two PDFs duplicates
                          (0–64; lower = stricter.  8 is a good default.)
    auto_delete         : if True, delete duplicates without prompting (CI/batch use)
    """

    try:
        import fitz  # noqa: F401  — just check availability
    except ImportError:
        print(f"\n{Fore.YELLOW}⚠  PyMuPDF not installed — skipping duplicate detection.")
        print(f"   Install with:  pip install pymupdf{Style.RESET_ALL}")
        return

    print(f"\n{Fore.CYAN}[5.5/7] Detecting duplicate PDFs...{Style.RESET_ALL}")

    unit_dirs = sorted(
        [d for d in base_dir.iterdir() if d.is_dir() and d.name.startswith("Unit_")],
        key=lambda p: _natural_sort_key(p)
    )

    total_removed = 0

    for unit_dir in unit_dirs:
        for res_id in resource_type_ids:
            resource_name = RESOURCE_TYPES.get(res_id, res_id)
            resource_dir  = unit_dir / resource_name

            if not resource_dir.exists():
                continue

            pdf_files = sorted(resource_dir.glob("*.pdf"), key=_natural_sort_key)
            if len(pdf_files) < 2:
                continue  # nothing to compare

            print(f"\n  {Fore.BLUE}{unit_dir.name} / {resource_name}{Style.RESET_ALL}"
                  f"  ({len(pdf_files)} PDFs)")

            # ── Step 1+2: size filter → fingerprint → find duplicates ──────
            # Pre-compute file sizes (just a stat() call, no file reading)
            file_sizes: Dict[Path, int] = {p: p.stat().st_size for p in pdf_files}

            # Group files by exact size — only same-size files can be duplicates
            size_groups: Dict[int, List[Path]] = {}
            for p in pdf_files:
                size_groups.setdefault(file_sizes[p], []).append(p)

            # Only fingerprint files that share an exact size with another file
            # Files with unique sizes are skipped entirely (no pHash computation)
            candidates: set = set()
            for size, group in size_groups.items():
                if len(group) > 1:
                    candidates.update(group)

            skipped = len(pdf_files) - len(candidates)
            if skipped:
                print(f"    {Fore.YELLOW}↷  Skipped {skipped} file(s) with unique sizes "
                      f"(no pHash needed){Style.RESET_ALL}")

            if not candidates:
                print(f"    {Fore.GREEN}✓ No duplicates found (all files have unique sizes){Style.RESET_ALL}")
                continue

            # Only fingerprint the candidate files (same-size pairs)
            print(f"    Hashing {len(candidates)} same-size file(s)...")
            fingerprints: Dict[Path, Optional[List[int]]] = {}
            for pdf in candidates:
                print(f"      {pdf.name} …", end="\r")
                fingerprints[pdf] = _pdf_fingerprint(
                    pdf,
                    sample_count=sample_count,
                    skip_edges=skip_edges
                )
            print(" " * 60, end="\r")  # clear progress line

            # Union-Find to cluster duplicates (only within same-size groups)
            parent = {p: p for p in pdf_files}

            def find(x):
                while parent[x] != x:
                    parent[x] = parent[parent[x]]
                    x = parent[x]
                return x

            def union(x, y):
                parent[find(x)] = find(y)

            # Only compare within same-size groups (never across different sizes)
            for size, group in size_groups.items():
                if len(group) < 2:
                    continue
                for i, fa in enumerate(group):
                    for fb in group[i + 1:]:
                        ha = fingerprints.get(fa)
                        hb = fingerprints.get(fb)
                        if ha is None or hb is None:
                            continue
                        if _are_duplicates(ha, hb, threshold=hash_threshold):
                            union(fa, fb)

            # Group by root
            groups: Dict[Path, List[Path]] = {}
            for p in pdf_files:
                root = find(p)
                groups.setdefault(root, []).append(p)

            # Only care about groups with more than one member
            dup_groups = [g for g in groups.values() if len(g) > 1]

            if not dup_groups:
                print(f"    {Fore.GREEN}✓ No duplicates found{Style.RESET_ALL}")
                continue

            # ── Step 3: report + action ────────────────────────────────────
            files_to_delete: List[Path] = []

            for group in dup_groups:
                group_sorted = sorted(group, key=_natural_sort_key)
                keep    = group_sorted[0]
                to_del  = group_sorted[1:]

                print(f"\n    {Fore.YELLOW}⚠  Duplicate group detected:{Style.RESET_ALL}")
                print(f"      KEEP   → {keep.name}")
                for d in to_del:
                    print(f"      DELETE → {d.name}")

                if auto_delete:
                    files_to_delete.extend(to_del)
                else:
                    print(f"\n    {Fore.CYAN}Action? "
                          f"[d] delete duplicates  "
                          f"[s] skip this group  "
                          f"[a] delete ALL remaining groups automatically"
                          f"{Style.RESET_ALL} ",
                          end="")
                    choice = input().strip().lower()

                    if choice == "d":
                        files_to_delete.extend(to_del)
                    elif choice == "a":
                        files_to_delete.extend(to_del)
                        auto_delete = True   # cascade to remaining groups
                    # "s" or anything else → skip

            # ── Step 4: delete ─────────────────────────────────────────────
            deleted_in_folder = 0
            for victim in files_to_delete:
                try:
                    victim.unlink()
                    print(f"    {Fore.RED}✗  Deleted:{Style.RESET_ALL} {victim.name}")
                    deleted_in_folder += 1
                    total_removed     += 1
                except Exception as e:
                    print(f"    {Fore.YELLOW}⚠  Could not delete {victim.name}: {e}{Style.RESET_ALL}")

            # ── Step 5: renumber survivors ─────────────────────────────────
            if deleted_in_folder > 0:
                survivors = sorted(resource_dir.glob("*.pdf"), key=_natural_sort_key)
                _renumber_files(survivors)
                print(f"    {Fore.GREEN}✓  Renumbered {len(survivors)} remaining files{Style.RESET_ALL}")

    # ── Final summary ──────────────────────────────────────────────────────────
    print(f"\n{Fore.GREEN}{'=' * 70}{Style.RESET_ALL}")
    if total_removed:
        print(f"{Fore.GREEN}✓  Duplicate cleanup: removed {total_removed} file(s){Style.RESET_ALL}")
    else:
        print(f"{Fore.GREEN}✓  Duplicate cleanup: no duplicates found across all folders{Style.RESET_ALL}")
    print(f"{Fore.GREEN}{'=' * 70}{Style.RESET_ALL}")


# ── Quick standalone test ──────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python pdf_dedup.py <downloads_folder> [resource_ids]")
        print("Example: python pdf_dedup.py downloads/UE23CS352A 2 6")
        sys.exit(1)

    folder     = Path(sys.argv[1])
    res_ids    = sys.argv[2:] if len(sys.argv) > 2 else list(RESOURCE_TYPES.keys())
    deduplicate_pdfs_in_folder(folder, res_ids)
