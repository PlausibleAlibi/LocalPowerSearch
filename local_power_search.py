#!/usr/bin/env python3
"""
local_power_search.py — Lightweight local document indexing and search tool.

Designed for recovered hard-drive triage: recursively scans a folder,
extracts partial text from common document types, stores metadata and
searchable text in SQLite FTS5, and exposes scan/search/show/stats commands.

Usage:
    python local_power_search.py scan /mnt/recovered
    python local_power_search.py search "bankruptcy"
    python local_power_search.py show 42
    python local_power_search.py stats
"""

import argparse
import csv
import hashlib
import logging
import re
import sqlite3
import sys
from collections.abc import Callable
from datetime import datetime, timezone
from email import policy as email_policy
from email.parser import BytesParser
from pathlib import Path

# ---------------------------------------------------------------------------
# Optional third-party imports — each wrapped so the tool still works if a
# library isn't installed; unsupported file types are simply skipped.
# ---------------------------------------------------------------------------

try:
    import pypdf
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    import docx  # python-docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_DB = "recovered_index.db"
DEFAULT_MAX_CHARS = 25_000
DEFAULT_MAX_PDF_PAGES = 10
DEFAULT_SEARCH_LIMIT = 20

# Encoding preference order used by all text extractors
DEFAULT_ENCODINGS: tuple[str, ...] = ("utf-8", "latin-1")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
)
log = logging.getLogger("local_power_search")

# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------


def open_db(db_path: str) -> sqlite3.Connection:
    """Open (or create) the SQLite database and ensure schema exists."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    _apply_schema(conn)
    return conn


def _apply_schema(conn: sqlite3.Connection) -> None:
    """Create tables if they don't exist. Safe to call repeatedly."""
    conn.executescript(
        """
        PRAGMA journal_mode=WAL;

        CREATE TABLE IF NOT EXISTS files (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            path         TEXT    UNIQUE NOT NULL,
            filename     TEXT    NOT NULL,
            extension    TEXT    NOT NULL,
            size_bytes   INTEGER NOT NULL,
            modified_utc TEXT    NOT NULL,
            sha256       TEXT,
            status       TEXT    NOT NULL,   -- indexed | skipped | failed
            error_message TEXT,
            text_preview TEXT,
            indexed_utc  TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS scan_stats (
            id           INTEGER PRIMARY KEY CHECK (id = 1),
            total_seen   INTEGER NOT NULL DEFAULT 0,
            last_scan_utc TEXT
        );

        CREATE VIRTUAL TABLE IF NOT EXISTS files_fts
            USING fts5(
                filename,
                path,
                text_preview,
                content='files',
                content_rowid='id'
            );

        -- keep FTS in sync with files rows
        CREATE TRIGGER IF NOT EXISTS files_ai AFTER INSERT ON files BEGIN
            INSERT INTO files_fts(rowid, filename, path, text_preview)
            VALUES (new.id, new.filename, new.path, new.text_preview);
        END;

        CREATE TRIGGER IF NOT EXISTS files_au AFTER UPDATE ON files BEGIN
            INSERT INTO files_fts(files_fts, rowid, filename, path, text_preview)
            VALUES ('delete', old.id, old.filename, old.path, old.text_preview);
            INSERT INTO files_fts(rowid, filename, path, text_preview)
            VALUES (new.id, new.filename, new.path, new.text_preview);
        END;

        CREATE TRIGGER IF NOT EXISTS files_ad AFTER DELETE ON files BEGIN
            INSERT INTO files_fts(files_fts, rowid, filename, path, text_preview)
            VALUES ('delete', old.id, old.filename, old.path, old.text_preview);
        END;
        """
    )
    conn.commit()


# ---------------------------------------------------------------------------
# File metadata helpers
# ---------------------------------------------------------------------------


def file_modified_utc(path: Path) -> str:
    """Return the file mtime as an ISO-8601 UTC string."""
    mtime = path.stat().st_mtime
    return datetime.fromtimestamp(mtime, tz=timezone.utc).isoformat()


def sha256_of_file(path: Path, chunk_size: int = 65536) -> str:
    """Compute SHA-256 of a file. Returns empty string on error."""
    h = hashlib.sha256()
    try:
        with open(path, "rb") as fh:
            while True:
                block = fh.read(chunk_size)
                if not block:
                    break
                h.update(block)
        return h.hexdigest()
    except OSError:
        return ""


def needs_reindex(conn: sqlite3.Connection, path: Path) -> bool:
    """Return True if the file is new or has changed since last index."""
    row = conn.execute(
        "SELECT size_bytes, modified_utc, status FROM files WHERE path = ?",
        (str(path),),
    ).fetchone()
    if row is None:
        return True  # never seen
    if row["status"] == "failed":
        return True  # retry failures
    current_size = path.stat().st_size
    current_mtime = file_modified_utc(path)
    return row["size_bytes"] != current_size or row["modified_utc"] != current_mtime


# ---------------------------------------------------------------------------
# Text extraction — one function per supported extension
# ---------------------------------------------------------------------------


def _truncate(text: str, max_chars: int) -> str:
    """Return at most max_chars characters of text."""
    return text[:max_chars] if len(text) > max_chars else text


def extract_text_txt(path: Path, max_chars: int) -> str:
    """Plain text: try UTF-8 then latin-1 fallback."""
    for enc in DEFAULT_ENCODINGS:
        try:
            with open(path, "r", encoding=enc, errors="replace") as fh:
                return _truncate(fh.read(max_chars), max_chars)
        except OSError:
            break
    return ""


def extract_text_pdf(path: Path, max_chars: int, max_pages: int) -> str:
    """PDF: read first max_pages pages via pypdf."""
    if not HAS_PYPDF:
        raise ImportError("pypdf not installed")
    chunks: list[str] = []
    total = 0
    with open(path, "rb") as fh:
        reader = pypdf.PdfReader(fh, strict=False)
        num_pages = len(reader.pages)
        for page_num in range(min(num_pages, max_pages)):
            if total >= max_chars:
                break
            try:
                page_text = reader.pages[page_num].extract_text() or ""
            except Exception:  # noqa: BLE001 — corrupt page, keep going
                continue
            remaining = max_chars - total
            chunks.append(page_text[:remaining])
            total += len(page_text)
    return "".join(chunks)


def extract_text_docx(path: Path, max_chars: int) -> str:
    """DOCX: join paragraph text."""
    if not HAS_DOCX:
        raise ImportError("python-docx not installed")
    doc = docx.Document(str(path))
    parts: list[str] = []
    total = 0
    for para in doc.paragraphs:
        if total >= max_chars:
            break
        text = para.text
        remaining = max_chars - total
        parts.append(text[:remaining])
        total += len(text)
    return "\n".join(parts)


def extract_text_csv(path: Path, max_chars: int) -> str:
    """CSV: read rows as space-joined tokens."""
    parts: list[str] = []
    total = 0
    for enc in DEFAULT_ENCODINGS:
        try:
            with open(path, "r", encoding=enc, errors="replace", newline="") as fh:
                reader = csv.reader(fh)
                for row in reader:
                    if total >= max_chars:
                        break
                    line = " ".join(cell.strip() for cell in row if cell.strip())
                    if not line:
                        continue
                    remaining = max_chars - total
                    parts.append(line[:remaining])
                    total += len(line)
            return "\n".join(parts)
        except (OSError, csv.Error):
            parts = []
            total = 0
    return "\n".join(parts)


def extract_text_xlsx(path: Path, max_chars: int) -> str:
    """XLSX: iterate sheets and cells (read_only mode for speed)."""
    if not HAS_OPENPYXL:
        raise ImportError("openpyxl not installed")
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    parts: list[str] = []
    total = 0
    for sheet in wb.worksheets:
        if total >= max_chars:
            break
        for row in sheet.iter_rows(values_only=True):
            if total >= max_chars:
                break
            tokens = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if not tokens:
                continue
            line = " ".join(tokens)
            remaining = max_chars - total
            parts.append(line[:remaining])
            total += len(line)
    wb.close()
    return "\n".join(parts)


def extract_text_html(path: Path, max_chars: int) -> str:
    """HTML/HTM: strip tags with BeautifulSoup, fallback to regex."""
    try:
        raw = path.read_bytes()
    except OSError:
        return ""
    if HAS_BS4:
        soup = BeautifulSoup(raw, "html.parser")
        text = soup.get_text(separator=" ", strip=True)
        return _truncate(text, max_chars)
    for enc in DEFAULT_ENCODINGS:
        try:
            decoded = raw.decode(enc, errors="replace")
            text = re.sub(r"<[^>]+>", " ", decoded)
            text = re.sub(r"\s+", " ", text).strip()
            return _truncate(text, max_chars)
        except (UnicodeDecodeError, ValueError):
            continue
    return ""


def extract_text_rtf(path: Path, max_chars: int) -> str:
    """RTF: naive strip of RTF control words; no external library needed."""
    for enc in DEFAULT_ENCODINGS:
        try:
            raw = path.read_text(encoding=enc, errors="replace")
            # Remove RTF control words and groups
            text = re.sub(r"\\\w+(-?\d+)?[ ]?", " ", raw)
            text = re.sub(r"[{}]", "", text)
            text = re.sub(r"\s+", " ", text).strip()
            return _truncate(text, max_chars)
        except OSError:
            break
    return ""


def extract_text_eml(path: Path, max_chars: int) -> str:
    """EML: extract plain-text body parts."""
    try:
        raw = path.read_bytes()
        msg = BytesParser(policy=email_policy.default).parsebytes(raw)
        parts: list[str] = []
        total = 0
        for part in msg.walk():
            if total >= max_chars:
                break
            if part.get_content_type() == "text/plain":
                try:
                    payload = part.get_content()
                except Exception:  # noqa: BLE001
                    payload = ""
                remaining = max_chars - total
                parts.append(payload[:remaining])
                total += len(payload)
        return "".join(parts)
    except Exception:  # noqa: BLE001
        return ""


# Map extensions to extractor callables.
# Each extractor receives (path, max_chars) or (path, max_chars, extra_arg).
EXTRACTORS: dict[str, tuple[str, Callable]] = {
    ".txt":  ("txt",  extract_text_txt),
    ".pdf":  ("pdf",  extract_text_pdf),
    ".docx": ("docx", extract_text_docx),
    ".csv":  ("csv",  extract_text_csv),
    ".xlsx": ("xlsx", extract_text_xlsx),
    ".html": ("html", extract_text_html),
    ".htm":  ("html", extract_text_html),
    ".rtf":  ("rtf",  extract_text_rtf),
    ".eml":  ("eml",  extract_text_eml),
}


def extract_text(path: Path, max_chars: int, max_pdf_pages: int) -> str:
    """Dispatch to the correct extractor. Raises on unsupported extension."""
    ext = path.suffix.lower()
    entry = EXTRACTORS.get(ext)
    if entry is None:
        raise ValueError(f"Unsupported extension: {ext}")
    kind, fn = entry
    if kind == "pdf":
        return fn(path, max_chars, max_pdf_pages)
    return fn(path, max_chars)


# ---------------------------------------------------------------------------
# Database write helpers
# ---------------------------------------------------------------------------


def upsert_file(
    conn: sqlite3.Connection,
    path: Path,
    status: str,
    text_preview: str = "",
    error_message: str = "",
    sha256: str = "",
) -> None:
    """Insert or update a row in the files table."""
    stat = path.stat()
    now = datetime.now(tz=timezone.utc).isoformat()
    conn.execute(
        """
        INSERT INTO files
            (path, filename, extension, size_bytes, modified_utc, sha256,
             status, error_message, text_preview, indexed_utc)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(path) DO UPDATE SET
            filename      = excluded.filename,
            extension     = excluded.extension,
            size_bytes    = excluded.size_bytes,
            modified_utc  = excluded.modified_utc,
            sha256        = excluded.sha256,
            status        = excluded.status,
            error_message = excluded.error_message,
            text_preview  = excluded.text_preview,
            indexed_utc   = excluded.indexed_utc
        """,
        (
            str(path),
            path.name,
            path.suffix.lower(),
            stat.st_size,
            file_modified_utc(path),
            sha256,
            status,
            error_message,
            text_preview,
            now,
        ),
    )
    conn.commit()


# ---------------------------------------------------------------------------
# scan command
# ---------------------------------------------------------------------------


def cmd_scan(args: argparse.Namespace) -> None:
    """Recursively scan a directory and index discovered files."""
    root = Path(args.root).resolve()
    if not root.is_dir():
        log.error("Root path is not a directory: %s", root)
        sys.exit(1)

    conn = open_db(args.db)
    max_chars: int = args.max_chars
    max_pdf_pages: int = args.max_pdf_pages
    compute_sha256: bool = args.sha256

    counters = {"seen": 0, "indexed": 0, "skipped": 0, "failed": 0, "unchanged": 0}

    log.info("Scanning %s", root)

    for file_path in sorted(root.rglob("*")):
        if not file_path.is_file():
            continue

        counters["seen"] += 1
        ext = file_path.suffix.lower()

        # Determine if we even support this extension
        if ext not in EXTRACTORS:
            if not needs_reindex(conn, file_path):
                counters["unchanged"] += 1
                continue
            upsert_file(conn, file_path, status="skipped")
            counters["skipped"] += 1
            log.debug("Skipped (unsupported): %s", file_path)
            continue

        # Skip unchanged files
        if not needs_reindex(conn, file_path):
            counters["unchanged"] += 1
            continue

        # Attempt extraction
        sha = sha256_of_file(file_path) if compute_sha256 else ""
        try:
            text_preview = extract_text(file_path, max_chars, max_pdf_pages)
            upsert_file(conn, file_path, status="indexed",
                        text_preview=text_preview, sha256=sha)
            counters["indexed"] += 1
            log.debug("Indexed: %s", file_path)
        except Exception as exc:  # noqa: BLE001 — resilience over crash
            err_msg = f"{type(exc).__name__}: {exc}"
            upsert_file(conn, file_path, status="failed",
                        error_message=err_msg, sha256=sha)
            counters["failed"] += 1
            log.warning("Failed %s — %s", file_path.name, err_msg)

    # Update total_seen
    conn.execute(
        """
        INSERT INTO scan_stats (id, total_seen, last_scan_utc)
        VALUES (1, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            total_seen = total_seen + excluded.total_seen,
            last_scan_utc = excluded.last_scan_utc
        """,
        (counters["seen"], datetime.now(tz=timezone.utc).isoformat()),
    )
    conn.commit()
    conn.close()

    log.info(
        "Scan complete — seen=%d indexed=%d skipped=%d failed=%d unchanged=%d",
        counters["seen"],
        counters["indexed"],
        counters["skipped"],
        counters["failed"],
        counters["unchanged"],
    )


# ---------------------------------------------------------------------------
# search command
# ---------------------------------------------------------------------------


def _snippet(text: str, query: str, window: int = 120) -> str:
    """Return a short snippet from text around the first query token match."""
    if not text:
        return ""
    # Find first occurrence of any query word (case-insensitive)
    for token in query.split():
        m = re.search(re.escape(token), text, re.IGNORECASE)
        if m:
            start = max(0, m.start() - window // 2)
            end = min(len(text), m.end() + window // 2)
            prefix = "…" if start > 0 else ""
            suffix = "…" if end < len(text) else ""
            return prefix + text[start:end].replace("\n", " ") + suffix
    return text[:window].replace("\n", " ")


def cmd_search(args: argparse.Namespace) -> None:
    """Search indexed files using FTS5 full-text search."""
    conn = open_db(args.db)
    query = args.query
    limit = args.limit

    # Build WHERE clauses for optional filters
    extra_where = ""
    params: list = [query, limit]

    if args.ext:
        norm_ext = args.ext if args.ext.startswith(".") else f".{args.ext}"
        extra_where += " AND f.extension = ?"
        params.insert(-1, norm_ext.lower())

    if args.path_contains:
        extra_where += " AND f.path LIKE ?"
        params.insert(-1, f"%{args.path_contains}%")

    sql = f"""
        SELECT
            f.id,
            f.filename,
            f.extension,
            f.modified_utc,
            f.path,
            f.text_preview,
            fts.rank
        FROM files_fts fts
        JOIN files f ON f.id = fts.rowid
        WHERE files_fts MATCH ?
          AND f.status = 'indexed'
          {extra_where}
        ORDER BY fts.rank
        LIMIT ?
    """

    rows = conn.execute(sql, params).fetchall()
    conn.close()

    if not rows:
        print("No results found.")
        return

    print(f"\nResults for: {query!r}  ({len(rows)} found)\n")
    print(f"{'ID':>5}  {'Ext':<7}  {'Modified':<24}  {'Filename'}")
    print("-" * 80)
    for row in rows:
        snippet = _snippet(row["text_preview"] or "", query)
        print(
            f"{row['id']:>5}  {row['extension']:<7}  {row['modified_utc']:<24}  "
            f"{row['filename']}"
        )
        print(f"       Path: {row['path']}")
        if snippet:
            print(f"       {snippet}")
        print()


# ---------------------------------------------------------------------------
# show command
# ---------------------------------------------------------------------------


def cmd_show(args: argparse.Namespace) -> None:
    """Display full metadata and text preview for a single file by ID."""
    conn = open_db(args.db)
    row = conn.execute(
        "SELECT * FROM files WHERE id = ?", (args.id,)
    ).fetchone()
    conn.close()

    if row is None:
        print(f"No file found with id={args.id}")
        return

    print()
    print(f"  ID          : {row['id']}")
    print(f"  Filename    : {row['filename']}")
    print(f"  Extension   : {row['extension']}")
    print(f"  Path        : {row['path']}")
    print(f"  Size        : {row['size_bytes']:,} bytes")
    print(f"  Modified    : {row['modified_utc']}")
    print(f"  SHA-256     : {row['sha256'] or '(not computed)'}")
    print(f"  Status      : {row['status']}")
    if row["error_message"]:
        print(f"  Error       : {row['error_message']}")
    print(f"  Indexed     : {row['indexed_utc']}")
    print()
    if row["text_preview"]:
        print("  --- Text preview ---")
        print(row["text_preview"][:2000])
        if len(row["text_preview"]) > 2000:
            print(f"  … [{len(row['text_preview'])} chars total]")
    else:
        print("  (no text preview)")
    print()


# ---------------------------------------------------------------------------
# stats command
# ---------------------------------------------------------------------------


def cmd_stats(args: argparse.Namespace) -> None:
    """Display summary statistics about the index."""
    conn = open_db(args.db)

    total_row = conn.execute(
        "SELECT total_seen, last_scan_utc FROM scan_stats WHERE id = 1"
    ).fetchone()
    total_seen = total_row["total_seen"] if total_row else 0
    last_scan = total_row["last_scan_utc"] if total_row else "never"

    status_rows = conn.execute(
        "SELECT status, COUNT(*) AS cnt FROM files GROUP BY status ORDER BY cnt DESC"
    ).fetchall()

    ext_rows = conn.execute(
        """
        SELECT extension, status, COUNT(*) AS cnt
        FROM files
        GROUP BY extension, status
        ORDER BY extension, status
        """
    ).fetchall()
    conn.close()

    print()
    print(f"  Database    : {args.db}")
    print(f"  Last scan   : {last_scan}")
    print(f"  Total seen  : {total_seen:,}")
    print()
    for sr in status_rows:
        print(f"  {sr['status']:<12}: {sr['cnt']:,}")
    print()
    print("  By extension and status:")
    print(f"  {'Extension':<12}  {'Status':<10}  {'Count':>7}")
    print("  " + "-" * 32)
    for er in ext_rows:
        print(f"  {er['extension'] or '(none)':<12}  {er['status']:<10}  {er['cnt']:>7}")
    print()


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="local_power_search",
        description="Lightweight local document indexing and search for hard-drive triage.",
    )
    parser.add_argument(
        "--db",
        default=DEFAULT_DB,
        help=f"Path to SQLite database (default: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Enable debug logging"
    )

    sub = parser.add_subparsers(dest="command", required=True)

    # --- scan ---
    sp = sub.add_parser("scan", help="Scan and index a directory")
    sp.add_argument("root", help="Root directory to scan")
    sp.add_argument(
        "--max-chars",
        type=int,
        default=DEFAULT_MAX_CHARS,
        help=f"Max characters extracted per file (default: {DEFAULT_MAX_CHARS})",
    )
    sp.add_argument(
        "--max-pdf-pages",
        type=int,
        default=DEFAULT_MAX_PDF_PAGES,
        help=f"Max PDF pages to read (default: {DEFAULT_MAX_PDF_PAGES})",
    )
    sp.add_argument(
        "--sha256",
        action="store_true",
        default=True,
        help="Compute SHA-256 hash for each file (default: enabled)",
    )
    sp.add_argument(
        "--no-sha256",
        dest="sha256",
        action="store_false",
        help="Skip SHA-256 computation",
    )

    # --- search ---
    sp = sub.add_parser("search", help="Search indexed files")
    sp.add_argument("query", help="Search query (FTS5 syntax supported)")
    sp.add_argument(
        "--limit",
        type=int,
        default=DEFAULT_SEARCH_LIMIT,
        help=f"Maximum results to return (default: {DEFAULT_SEARCH_LIMIT})",
    )
    sp.add_argument("--ext", help="Filter by file extension (e.g. pdf or .pdf)")
    sp.add_argument("--path-contains", help="Filter: path must contain this string")

    # --- show ---
    sp = sub.add_parser("show", help="Show metadata for a file by ID")
    sp.add_argument("id", type=int, help="File ID from the index")

    # --- stats ---
    sub.add_parser("stats", help="Show index statistics")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.verbose:
        log.setLevel(logging.DEBUG)
        logging.getLogger().setLevel(logging.DEBUG)

    dispatch = {
        "scan": cmd_scan,
        "search": cmd_search,
        "show": cmd_show,
        "stats": cmd_stats,
    }
    dispatch[args.command](args)


if __name__ == "__main__":
    main()
