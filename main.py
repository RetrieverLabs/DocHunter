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


# ============================================================
# CONFIG
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".txt",
    ".csv",
    ".json",
    ".xml",
    ".doc"
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
    line_number: int
    context: str


# ============================================================
# KEYWORDS
# ============================================================

def load_keywords(arg: str) -> List[str]:
    """
    Accepts:
    - raw keyword string
    - file containing keywords
    """
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
# TEXT EXTRACTION
# ============================================================

def extract_file_text(path: Path) -> List[str]:
    """
    Convert any supported file into list of text lines.
    """

    ext = path.suffix.lower()

    try:
        # ---------------- TXT / CSV ----------------
        if ext in {".txt", ".csv"}:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.readlines()

        # ---------------- DOC (legacy Word) ----------------
        if ext == ".doc":
            raw = textract.process(str(path))
            text = raw.decode("utf-8", errors="ignore")
            return text.splitlines()

        # ---------------- JSON ----------------
        if ext == ".json":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                data = json.load(f)
            return json.dumps(data, indent=2).splitlines()

        # ---------------- XML ----------------
        if ext == ".xml":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.readlines()

    except Exception:
        return []

    return []


# ============================================================
# SEARCH WORKER
# ============================================================

def search_file(args: Tuple[str, List[str]]) -> List[MatchResult]:
    file_path, keywords = args
    path = Path(file_path)

    lines = extract_file_text(path)
    results = []

    for i, line in enumerate(lines, start=1):
        for kw in keywords:
            if kw.lower() in line.lower():
                results.append(
                    MatchResult(
                        file_path=str(path),
                        extension=path.suffix,
                        keyword=kw,
                        line_number=i,
                        context=line.strip()[:300]
                    )
                )

    return results


# ============================================================
# PROGRESS TRACKER
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
# SCANNER
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
# CSV OUTPUT
# ============================================================

def write_csv(results: List[MatchResult], output="results.csv"):
    with open(output, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["File", "Extension", "Keyword", "Line", "Context"])

        for r in results:
            w.writerow([
                r.file_path,
                r.extension,
                r.keyword,
                r.line_number,
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
    print()

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
