import os
import sys
import csv
import time
import json
import textract
import multiprocessing as mp

from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple
from concurrent.futures import ProcessPoolExecutor, as_completed

from docx import Document
import fitz  # PyMuPDF


# ============================================================
# CONFIG
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".txt",
    ".csv",
    ".json",
    ".xml",
    ".doc",
    ".docx",
    ".pdf"
}

DEFAULT_WORKERS = max(1, (os.cpu_count() or 4) // 2)


# ============================================================
# DATA MODEL
# ============================================================

@dataclass
class MatchResult:
    file_path: str
    extension: str
    keyword: str
    location: str
    context: str


# ============================================================
# KEYWORDS
# ============================================================

def load_keywords(arg: str) -> List[str]:
    p = Path(arg)

    if p.exists() and p.is_file():
        with open(p, "r", encoding="utf-8", errors="ignore") as f:
            return [line.strip() for line in f if line.strip()]

    return [arg]


# ============================================================
# FILE DISCOVERY
# ============================================================

def find_files(root: Path) -> List[Path]:
    files = []
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(p)
    return files


# ============================================================
# TEXT EXTRACTORS
# ============================================================

def extract_txt(path: Path) -> List[str]:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.readlines()


def extract_json(path: Path) -> List[str]:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        data = json.load(f)
    return json.dumps(data, indent=2).splitlines()


def extract_xml(path: Path) -> List[str]:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.readlines()


def extract_doc(path: Path) -> List[str]:
    raw = textract.process(str(path))
    text = raw.decode("utf-8", errors="ignore")
    return text.splitlines()


def extract_docx(path: Path) -> List[str]:
    doc = Document(path)
    lines = []

    for p in doc.paragraphs:
        lines.append(p.text)

    return lines


def extract_pdf(path: Path) -> List[str]:
    doc = fitz.open(path)
    lines = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        for line in text.splitlines():
            lines.append(f"[Page {page_num+1}] {line}")

    return lines


# ============================================================
# DISPATCH TABLE
# ============================================================

EXTRACTORS = {
    ".txt": extract_txt,
    ".csv": extract_txt,
    ".json": extract_json,
    ".xml": extract_xml,
    ".doc": extract_doc,
    ".docx": extract_docx,
    ".pdf": extract_pdf,
}


# ============================================================
# SEARCH WORKER
# ============================================================

def search_file(args: Tuple[str, List[str]]) -> List[MatchResult]:
    file_path, keywords = args
    path = Path(file_path)

    extractor = EXTRACTORS.get(path.suffix.lower())

    if not extractor:
        return []

    try:
        lines = extractor(path)
    except Exception:
        return []

    results = []

    for i, line in enumerate(lines, start=1):
        line_l = line.lower()

        for kw in keywords:
            if kw.lower() in line_l:
                results.append(
                    MatchResult(
                        file_path=str(path),
                        extension=path.suffix,
                        keyword=kw,
                        location=f"Line {i}",
                        context=line.strip()[:300]
                    )
                )

    return results


# ============================================================
# PROGRESS
# ============================================================

class Progress:
    def __init__(self, total: int):
        self.total = total
        self.done = 0
        self.matches = 0

    def update(self, match_count: int):
        self.done += 1
        self.matches += match_count
        self.render()

    def render(self):
        if self.total == 0:
            return

        pct = (self.done / self.total) * 100
        bar_len = 30
        filled = int(bar_len * self.done / self.total)

        bar = "█" * filled + "░" * (bar_len - filled)

        print(
            f"\r[{bar}] {pct:5.1f}% "
            f"{self.done}/{self.total} files "
            f"| matches: {self.matches}",
            end=""
        )


# ============================================================
# SCAN ENGINE
# ============================================================

def scan(root: Path, keywords: List[str], workers: int) -> List[MatchResult]:
    files = find_files(root)
    progress = Progress(len(files))

    results: List[MatchResult] = []

    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = [
            executor.submit(search_file, (str(f), keywords))
            for f in files
        ]

        for fut in as_completed(futures):
            try:
                res = fut.result()
                results.extend(res)
                progress.update(len(res))
            except Exception:
                progress.update(0)

    print()
    return results


# ============================================================
# OUTPUT
# ============================================================

def write_csv(results: List[MatchResult], output="results.csv"):
    with open(output, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["File", "Extension", "Keyword", "Location", "Context"])

        for r in results:
            w.writerow([
                r.file_path,
                r.extension,
                r.keyword,
                r.location,
                r.context
            ])


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 3:
        print("Usage:")
        print("  python scan.py <folder> <keyword | keyword_file>")
        sys.exit(1)

    root = Path(sys.argv[1])
    keyword_arg = sys.argv[2]

    if not root.exists():
        print("Folder not found:", root)
        sys.exit(1)

    keywords = load_keywords(keyword_arg)

    print("\nDocHunter v0.1")
    print("====================")
    print("Folder   :", root)
    print("Keywords :", len(keywords))
    print("Workers  :", DEFAULT_WORKERS)
    print("Files    : scanning...\n")

    start = time.time()

    results = scan(root, keywords, DEFAULT_WORKERS)

    write_csv(results)

    elapsed = time.time() - start

    print("\nScan complete")
    print("Matches :", len(results))
    print(f"Time    : {elapsed:.2f}s")
    print("Output  : results.csv")


if __name__ == "__main__":
    mp.set_start_method("spawn", force=True)
    main()
