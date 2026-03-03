import json
import re
import shutil
import zipfile
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set


# ============================================================
# CONFIG
# ============================================================
BASE_DIR = Path(__file__).resolve().parent        # iDownload
ROOT = BASE_DIR / "AO3 Epub:Mobi"                # run INBOX is epub/
OUT_ROOT = ROOT / "AO3 main"                     # organized tree under epub/

MEMORY_DIR = BASE_DIR / "AO3_sort_memory"
MEMORY_DIR.mkdir(parents=True, exist_ok=True)

MEMORY_PATH = MEMORY_DIR / "ao3_memory.json"
PROGRESS_PATH = MEMORY_DIR / "ao3_progress.json"


def run_log_path_for_now() -> Path:
    ts = datetime.now().strftime("%Y%m%dT%H%M%S")
    return MEMORY_DIR / f"run_{ts}.jsonl"


RUN_LOG_PATH = run_log_path_for_now()
REL_LOG_PATH = MEMORY_DIR / "relationship_log.jsonl"

SUPPORTED_FILE_EXTS = {".epub", ".mobi", ".azw3", ".zip"}
TMP_EXTRACT_DIRNAME = "_tmp_zip_extract"

SUMMARY_MAX_CHARS = 800


# ============================================================
# Memory + Progress
# ============================================================
def default_memory() -> dict:
    return {
        "version": 1,
        "fandoms": {},                  # canonical -> {"aliases":[...]}
        "relationships_by_fandom": {},  # fandom canonical -> [rel_folder,...]
        "pairing_combo_rules": {},      # "fandom||pairA||pairB" -> rel_folder
        "fics": [],                     # list of tracked fic entries
    }


def default_progress() -> dict:
    return {"version": 1, "done": {}}


def load_json(path: Path, default_obj):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default_obj
    return default_obj


def save_json(path: Path, obj) -> None:
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def append_jsonl(path: Path, entry: dict) -> None:
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def load_memory() -> dict:
    mem = load_json(MEMORY_PATH, default_memory())
    base = default_memory()
    if isinstance(mem, dict):
        base.update(mem)
    base.setdefault("fandoms", {})
    base.setdefault("relationships_by_fandom", {})
    base.setdefault("pairing_combo_rules", {})
    base.setdefault("fics", [])
    return base


def save_memory(mem: dict) -> None:
    save_json(MEMORY_PATH, mem)


def load_progress() -> dict:
    prog = load_json(PROGRESS_PATH, default_progress())
    if not isinstance(prog, dict):
        return default_progress()
    prog.setdefault("done", {})
    return prog


def save_progress(prog: dict) -> None:
    save_json(PROGRESS_PATH, prog)


def item_key(p: Path) -> str:
    st = p.stat()
    if p.is_dir():
        return f"DIR|{p.name}|{int(st.st_mtime)}"
    return f"FILE|{p.name}|{st.st_size}|{int(st.st_mtime)}"


# ============================================================
# Helpers
# ============================================================
def normalize_alias(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip()).lower()


def sanitize_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else "untitled"


def has_chinese_chars(s: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", s))


def underscore_to_space_if_allowed(path: Path) -> Path:
    # Do NOT rename Chinese-titled items; else replace "_" with space
    if has_chinese_chars(path.name):
        return path
    if "_" not in path.name:
        return path
    new_name = path.name.replace("_", " ")
    new_path = path.with_name(new_name)
    if new_path.exists():
        return path
    path.rename(new_path)
    print(f"[Rename _-> ] {path.name} -> {new_path.name}")
    return new_path


def ensure_folder(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p


def remove_path(p: Path) -> None:
    if p.is_dir():
        shutil.rmtree(p)
    else:
        p.unlink()


# ============================================================
# Logging
# ============================================================
def log_action(run_id: str, action: str, src: str, dest_folder: str, note: str) -> None:
    append_jsonl(
        RUN_LOG_PATH,
        {
            "run_id": run_id,
            "when": datetime.now().isoformat(timespec="seconds"),
            "action": action,
            "src": src,
            "dest_folder": dest_folder,
            "note": note,
        },
    )


# ============================================================
# Minimal HTML -> text (EPUB internal HTML)
# ============================================================
def html_to_text(html: str) -> str:
    html = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html)
    html = re.sub(r"(?i)<br\s*/?>", "\n", html)
    html = re.sub(r"(?i)</p\s*>", "\n", html)
    html = re.sub(r"(?i)</div\s*>", "\n", html)
    text = re.sub(r"(?s)<.*?>", " ", html)
    text = (
        text.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&#39;", "'")
    )
    text = re.sub(r"[ \t\r]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


# ============================================================
# AO3 label parsing helpers
# ============================================================
def extract_block_after_label(text: str, label: str, stop_labels: List[str]) -> str:
    """
    Captures text that begins on the same line as 'label' or subsequent lines,
    until a stop label is hit.
    """
    lines = text.splitlines()
    out = []
    capturing = False
    label_low = label.lower()
    stop_low = [s.lower() for s in stop_labels]

    for ln in lines:
        s = ln.strip()
        low = s.lower()

        if low.startswith(label_low):
            capturing = True
            out.append(s[len(label) :].strip())
            continue

        if capturing:
            if any(low.startswith(x) for x in stop_low):
                break
            if s:
                out.append(s)

    return " ".join(out).strip()


def extract_fandom_block(text: str) -> str:
    return extract_block_after_label(
        text,
        "Fandom:",
        stop_labels=[
            "Relationship:",
            "Relationships:",
            "Characters:",
            "Language:",
            "Rating:",
            "Archive Warning:",
            "Category:",
            "Stats:",
            "Summary:",
            "Series:",
            "Additional Tags:",
        ],
    )


def extract_relationship_block(text: str) -> str:
    block = extract_block_after_label(
        text,
        "Relationship:",
        stop_labels=[
            "Characters:",
            "Language:",
            "Rating:",
            "Archive Warning:",
            "Category:",
            "Stats:",
            "Summary:",
            "Series:",
            "Fandom:",
            "Additional Tags:",
        ],
    )
    if block:
        return block
    return extract_block_after_label(
        text,
        "Relationships:",
        stop_labels=[
            "Characters:",
            "Language:",
            "Rating:",
            "Archive Warning:",
            "Category:",
            "Stats:",
            "Summary:",
            "Series:",
            "Fandom:",
            "Additional Tags:",
        ],
    )


def extract_summary_block(text: str) -> str:
    """
    Extracts AO3 Summary section.
    AO3 often has a line 'Summary' then the summary text.
    Stops when it hits a new header.
    """
    lines = [ln.rstrip() for ln in text.splitlines()]
    out: List[str] = []
    capturing = False

    stop_markers = [
        "Notes",
        "Chapter",
        "Chapter Summary",
        "Chapter Notes",
        "End Notes",
        "Work Text",
        "Inspired by",
        "Stats:",
        "Language:",
        "Relationships:",
        "Relationship:",
        "Characters:",
        "Additional Tags:",
        "Fandom:",
        "Rating:",
        "Archive Warning:",
        "Categories:",
        "Category:",
        "Series:",
    ]

    for ln in lines:
        s = ln.strip()

        if not capturing:
            if s.lower() == "summary":
                capturing = True
            continue

        if not s:
            out.append("")
            continue

        if any(s.startswith(m) for m in stop_markers):
            break

        out.append(s)

    summary = "\n".join(out).strip()
    summary = re.sub(r"\n{3,}", "\n\n", summary)
    return summary


def extract_author(text: str) -> Optional[str]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines[:180]:
        m = re.match(r"^by\s+(.+)$", ln, flags=re.IGNORECASE)
        if m:
            a = m.group(1).strip()
            if a:
                return a
    return None


def extract_updated_date(text: str) -> Optional[date]:
    m = re.search(r"\bUpdated:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})\b", text)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d").date()
    except Exception:
        return None


def extract_series_entries(text: str) -> List[Tuple[str, Optional[int]]]:
    block = extract_block_after_label(
        text,
        "Series:",
        stop_labels=[
            "Relationship:",
            "Relationships:",
            "Characters:",
            "Language:",
            "Rating:",
            "Archive Warning:",
            "Category:",
            "Stats:",
            "Summary:",
            "Fandom:",
            "Additional Tags:",
        ],
    )
    if not block:
        return []

    raw_parts = re.split(r"[|;]+|(?:\s*/\s*)|(?:\s{2,})|,(?![^()]*\))", block)
    raw_parts = [p.strip() for p in raw_parts if p.strip()]

    out: List[Tuple[str, Optional[int]]] = []
    for part in raw_parts:
        p = part.strip()
        m1 = re.match(r"^part\s+([0-9]+)\s+of\s+(.+)$", p, flags=re.IGNORECASE)
        if m1:
            pn = int(m1.group(1))
            title = m1.group(2).strip()
            title = re.sub(
                r"\(\s*part\s+\d+\s+of\s+\d+\s*\)\s*$", "", title, flags=re.IGNORECASE
            ).strip()
            title = re.sub(r"^[•\-–—]+\s*", "", title).strip()
            if title:
                out.append((title, pn))
            continue

        m2 = re.search(
            r"\(\s*part\s+([0-9]+)\s+of\s+[0-9]+\s*\)\s*$", p, flags=re.IGNORECASE
        )
        if m2:
            pn = int(m2.group(1))
            title = re.sub(
                r"\(\s*part\s+\d+\s+of\s+\d+\s*\)\s*$", "", p, flags=re.IGNORECASE
            ).strip()
            title = re.sub(r"^[•\-–—]+\s*", "", title).strip()
            if title:
                out.append((title, pn))
            continue

        title_only = re.sub(r"^[•\-–—]+\s*", "", p).strip()
        if title_only:
            out.append((title_only, None))

    seen = set()
    deduped = []
    for t, pn in out:
        key = (t, pn)
        if key not in seen:
            seen.add(key)
            deduped.append((t, pn))
    return deduped


# ============================================================
# Relationship parsing: relationship folders = any tag containing "/"
# ============================================================
def split_relationship_tags(rel_block: str) -> List[str]:
    if not rel_block:
        return []
    parts = re.split(r"[|,;]+", rel_block)
    return [p.strip() for p in parts if p.strip()]


def extract_slash_pairs(text: str) -> List[str]:
    rel_block = extract_relationship_block(text)
    tags = split_relationship_tags(rel_block)
    pairs = []
    for t in tags:
        t2 = re.sub(r"\s+", " ", t).strip()
        if "/" in t2:
            pairs.append(t2)

    seen = set()
    out = []
    for p in pairs:
        k = p.lower()
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out


# ============================================================
# Extract text (open once, never printed)
# ============================================================
def extract_text_from_epub(epub_path: Path, max_files: int = 6) -> str:
    try:
        with zipfile.ZipFile(epub_path, "r") as z:
            names = z.namelist()
            htmls = [
                n for n in names if n.lower().endswith((".xhtml", ".html", ".htm"))
            ]
            htmls.sort()
            combined = []
            for n in htmls[:max_files]:
                data = z.read(n)
                try:
                    s = data.decode("utf-8", errors="ignore")
                except Exception:
                    s = data.decode(errors="ignore")
                combined.append(html_to_text(s))
            return "\n".join(combined)
    except Exception:
        return ""


def extract_text_from_zip(zip_path: Path) -> str:
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            names = z.namelist()
            epubs = [n for n in names if n.lower().endswith(".epub")]
            if epubs:
                tmp_dir = ensure_folder(OUT_ROOT / TMP_EXTRACT_DIRNAME)
                z.extract(epubs[0], path=tmp_dir)
                extracted = tmp_dir / epubs[0]
                return extract_text_from_epub(extracted)
            htmls = [
                n for n in names if n.lower().endswith((".xhtml", ".html", ".htm"))
            ]
            htmls.sort()
            combined = []
            for n in htmls[:6]:
                data = z.read(n)
                try:
                    s = data.decode("utf-8", errors="ignore")
                except Exception:
                    s = data.decode(errors="ignore")
                combined.append(html_to_text(s))
            return "\n".join(combined)
    except Exception:
        return ""


def extract_text_for_item(item: Path) -> str:
    if item.is_dir():
        # find first epub/zip inside
        for p in sorted(item.rglob("*")):
            if p.is_file() and p.suffix.lower() in {".epub", ".zip"}:
                return extract_text_for_item(p)
        return ""
    ext = item.suffix.lower()
    if ext == ".epub":
        return extract_text_from_epub(item)
    if ext == ".zip":
        return extract_text_from_zip(item)
    return ""  # mobi/azw3 not parsed


# ============================================================
# Series/author subfolder (auto)
# ============================================================
def choose_series_or_author_subfolder(text: str, author: str) -> Optional[str]:
    series_entries = extract_series_entries(text) if text else []
    if len(series_entries) == 1:
        return sanitize_name(series_entries[0][0])
    if len(series_entries) >= 2:
        return sanitize_name(author or "unknown")
    return None


# ============================================================
# Newest logic + conflict naming
# ============================================================
def get_item_date(text: str, path: Path) -> date:
    upd = extract_updated_date(text) if text else None
    if upd:
        return upd
    return datetime.fromtimestamp(path.stat().st_mtime).date()


def by_author_name(original_name: str, author: str) -> str:
    base = Path(original_name).stem
    suf = Path(original_name).suffix
    return f"{base} by {sanitize_name(author or 'unknown')}{suf}"


def existing_is_older(existing_path: Path, incoming_date: date) -> bool:
    existing_date = datetime.fromtimestamp(existing_path.stat().st_mtime).date()
    return incoming_date > existing_date


# ============================================================
# Existing index under main/ for duplicate routing
# ============================================================
def build_existing_index() -> Dict[str, List[Path]]:
    idx: Dict[str, List[Path]] = {}
    if not OUT_ROOT.exists():
        return idx
    for p in OUT_ROOT.rglob("*"):
        if p.is_file() and p.suffix.lower() in SUPPORTED_FILE_EXTS:
            idx.setdefault(p.name.lower(), []).append(p)
    return idx


def pick_best_existing_match(paths: List[Path]) -> Path:
    return max(paths, key=lambda x: x.stat().st_mtime)


def find_existing_destination(
    existing_idx: Dict[str, List[Path]], incoming_name: str, author: str
) -> Optional[Path]:
    exact = existing_idx.get(incoming_name.lower(), [])
    if exact:
        return pick_best_existing_match(exact)
    by_name = by_author_name(incoming_name, author)
    by_matches = existing_idx.get(by_name.lower(), [])
    if by_matches:
        return pick_best_existing_match(by_matches)
    return None


# ============================================================
# Fandom prompts + memory (DISPLAY ONLY fandom tags)
# ============================================================
def find_matching_fandom(fandom_text: str, mem: dict) -> Optional[str]:
    if not fandom_text:
        return None
    hay = normalize_alias(fandom_text)
    for canonical, obj in mem["fandoms"].items():
        for a in obj.get("aliases", []):
            needle = normalize_alias(a)
            if needle and needle in hay:
                return canonical
    return None


def suggest_aliases(fandom_text: str) -> List[str]:
    if not fandom_text:
        return []
    parts = re.split(r"[|,;/]+", fandom_text)
    parts = [p.strip() for p in parts if p.strip()]
    candidates = []
    if fandom_text.strip():
        candidates.append(fandom_text.strip())
    for p in parts:
        if p not in candidates:
            candidates.append(p)
    return candidates[:20]


def update_fandom_mem(mem: dict, canonical: str, fandom_raw: str) -> None:
    mem["fandoms"].setdefault(canonical, {"aliases": []})
    aliases = mem["fandoms"][canonical].setdefault("aliases", [])
    existing_norm = {normalize_alias(a) for a in aliases}
    for a in suggest_aliases(fandom_raw):
        if normalize_alias(a) not in existing_norm:
            aliases.append(a)
            existing_norm.add(normalize_alias(a))


def prompt_for_fandom(item_name: str, fandom_raw: str, mem: dict) -> str:
    existing = sorted(mem["fandoms"].keys(), key=lambda x: x.lower())

    print("\n" + "=" * 80)
    print(f"ITEM: {item_name}")
    print(f"Fandom tag(s): {fandom_raw if fandom_raw else '[none found]'}")
    print("-" * 80)

    if existing:
        print("Existing fandom folders:")
        for i, c in enumerate(existing, start=1):
            print(f"  {i}. {c}")
    else:
        print("No fandoms in memory yet.")

    print("\nChoose fandom:")
    print("  [number] -> existing fandom")
    print("  n        -> create NEW fandom")
    print("  u        -> undecided (leave item in inbox)")
    print("  s        -> skip item this run")

    while True:
        c = input("\nYour choice: ").strip().lower()
        if c == "s":
            return "__SKIP__"
        if c == "u":
            return "__UNDECIDED__"
        if c == "n":
            new_name = sanitize_name(input("New fandom folder name: ").strip())
            if new_name:
                return new_name
            print("Name cannot be empty.")
            continue
        if c.isdigit():
            idx = int(c)
            if 1 <= idx <= len(existing):
                return existing[idx - 1]
            print("Number out of range.")
            continue
        print("Invalid input.")


# ============================================================
# Relationship prompts + memory (DISPLAY relationships + summary)
# ============================================================
def remember_relationship(mem: dict, fandom: str, rel_folder: str) -> None:
    mem["relationships_by_fandom"].setdefault(fandom, [])
    if rel_folder not in mem["relationships_by_fandom"][fandom]:
        mem["relationships_by_fandom"][fandom].append(rel_folder)


def combo_key(fandom: str, pairs: List[str]) -> str:
    norm = sorted(normalize_alias(p) for p in pairs)
    return f"{fandom}||" + "||".join(norm)


def _truncate(s: str, max_chars: int) -> str:
    s = (s or "").strip()
    if len(s) <= max_chars:
        return s
    return s[:max_chars].rstrip() + " ..."


def prompt_relationship_choice(
    item_name: str,
    fandom: str,
    slash_pairs: List[str],
    mem: dict,
    rel_raw: str,
    summary: str,
) -> str:
    existing = mem["relationships_by_fandom"].get(fandom, [])

    print("\n" + "-" * 80)
    print(f"RELATIONSHIP STEP for: {item_name}")
    print(f"Fandom folder: {fandom}")
    print(f"Relationship tag(s): {rel_raw if rel_raw else '[none found]'}")

    if summary:
        print("\nSummary:")
        print(_truncate(summary, SUMMARY_MAX_CHARS))
    else:
        print("\nSummary: [not found]")

    if slash_pairs:
        print("\nSlash-pairs found (/):")
        for i, p in enumerate(slash_pairs, start=1):
            print(f"  {i}. {p}")

    if existing:
        print("\nKnown relationship folders for this fandom:")
        for i, r in enumerate(existing, start=1):
            print(f"  {i}. {r}")
    else:
        print("\nNo known relationship folders for this fandom yet.")

    print("\nChoose destination level:")
    print("  f -> keep in <fandom> (undecided relationship)")
    print("  [number] -> choose existing relationship folder")
    print("  p -> use a slash-pair as relationship folder (pick from detected pairs)")
    print("  n -> create NEW relationship folder")
    print("  s -> skip item")

    while True:
        c = input("\nYour choice: ").strip().lower()
        if c == "s":
            return "__SKIP__"
        if c == "f":
            return "__FANDOM_LEVEL__"
        if c == "n":
            new_rel = sanitize_name(input("New relationship folder name: ").strip())
            if new_rel:
                return new_rel
            print("Name cannot be empty.")
            continue
        if c == "p":
            if not slash_pairs:
                print("No slash-pairs detected for this item.")
                continue
            pick = input(f"Pick slash-pair 1-{len(slash_pairs)}: ").strip()
            if pick.isdigit():
                idx = int(pick)
                if 1 <= idx <= len(slash_pairs):
                    return sanitize_name(slash_pairs[idx - 1])
            print("Invalid pair selection.")
            continue
        if c.isdigit():
            idx = int(c)
            if 1 <= idx <= len(existing):
                return existing[idx - 1]
            print("Number out of range.")
            continue
        print("Invalid input.")


def resolve_relationship(
    item_name: str,
    fandom: str,
    slash_pairs: List[str],
    mem: dict,
    rel_raw: str,
    summary: str,
) -> Optional[str]:
    if len(slash_pairs) == 0:
        return None

    # One slash-pair: auto if known; else ask
    if len(slash_pairs) == 1:
        candidate = sanitize_name(slash_pairs[0])
        existing = mem["relationships_by_fandom"].get(fandom, [])
        if candidate in existing:
            return candidate

        choice = prompt_relationship_choice(
            item_name, fandom, slash_pairs, mem, rel_raw, summary
        )
        if choice == "__SKIP__":
            return "__SKIP__"
        if choice == "__FANDOM_LEVEL__":
            return None
        return choice

    # Multiple slash-pairs: combo rule or ask and save
    key = combo_key(fandom, slash_pairs)
    if key in mem["pairing_combo_rules"]:
        return mem["pairing_combo_rules"][key]

    choice = prompt_relationship_choice(
        item_name, fandom, slash_pairs, mem, rel_raw, summary
    )
    if choice == "__SKIP__":
        return "__SKIP__"
    if choice == "__FANDOM_LEVEL__":
        return None

    mem["pairing_combo_rules"][key] = choice
    return choice


# ============================================================
# Item info (extract once)
# ============================================================
@dataclass
class ItemInfo:
    path: Path
    text: str
    fandomraw: str
    relraw: str
    summary: str
    slashpairs: List[str]
    author: str
    itemdate: date
    title: str
    ao3_url: Optional[str]
    chapters_local: Optional[int]
    complete: bool


def extract_item_info(p: Path) -> ItemInfo:
    p2 = underscore_to_space_if_allowed(p)
    text = extract_text_for_item(p2)

    fandom_raw = extract_fandom_block(text) if text else ""
    rel_raw = extract_relationship_block(text) if text else ""
    summary = extract_summary_block(text) if text else ""
    slash_pairs = extract_slash_pairs(text) if text else []
    author = extract_author(text) if text else None
    author = author or "unknown"
    item_date = get_item_date(text, p2)

    # log relationship tags and summary for future mapping; safe (does NOT print to terminal)
    append_jsonl(
        REL_LOG_PATH,
        {
            "when": datetime.now().isoformat(timespec="seconds"),
            "item": p2.name,
            "format": "folder" if p2.isdir() else p2.suffix.lower(),
            "fandom_raw": fandom_raw,
            "relationship_raw": rel_raw,
            "slashpairs": slash_pairs,
            "summary": _truncate(summary, 2000),  # keep log manageable
            "author_guess": author,
            "updated_guess": str(extract_updated_date(text) if text else None),
        },
    )

    # --- NEW: derive title, ao3_url, chapters_local, complete ---
    base_title = p2.stem
    title = base_title

    # AO3 URL: look for works/<id> in the text
    ao3_url = None
    if text:
        m_url = re.search(r"https?://archiveofourown\.org/works/(\d+)", text)
        if m_url:
            ao3_url = f"https://archiveofourown.org/works/{m_url.group(1)}"

    # Chapters: parse "Chapters: X/?" or "Chapters: X of Y" style lines
    chapters_local = None
    complete = False
    if text:
        m_ch = re.search(r"Chapters:\s*(\d+)\s*/\s*([0-9?]+)", text)
        if not m_ch:
            m_ch = re.search(r"Chapter[s]?:\s*(\d+)(?:\s*of\s*([0-9?]+))?", text)
        if m_ch:
            try:
                chapters_local = int(m_ch.group(1))
            except Exception:
                chapters_local = None
            total_raw = (
                m_ch.group(2) if m_ch.lastindex and m_ch.group(2) is not None else None
            )
            if total_raw and total_raw != "?":
                try:
                    total_int = int(total_raw)
                    if chapters_local is not None and chapters_local >= total_int:
                        complete = True
                except Exception:
                    complete = False

    return ItemInfo(
        path=p2,
        text=text,
        fandomraw=fandom_raw,
        relraw=rel_raw,
        summary=summary,
        slashpairs=slash_pairs,
        author=author,
        itemdate=item_date,
        title=title,
        ao3_url=ao3_url,
        chapters_local=chapters_local,
        complete=complete,
    )


# ============================================================
# Inbox iteration
# ============================================================
def iter_inbox_items() -> List[Path]:
    items = []
    for p in ROOT.iterdir():
        if p.name in {
            OUT_ROOT.name,
            MEMORY_PATH.name,
            PROGRESS_PATH.name,
            RUN_LOG_PATH.name,
            REL_LOG_PATH.name,
            ".venv",
            TMP_EXTRACT_DIRNAME,
        }:
            continue
        if p.is_file() and p.suffix.lower() not in SUPPORTED_FILE_EXTS:
            continue
        items.append(p)
    return sorted(items, key=lambda x: x.name.lower())


# ============================================================
# MAIN
# ============================================================
def main():
    run_id = datetime.now().strftime("%Y%m%dT%H%M%S")
    run_when = datetime.now().isoformat(timespec="seconds")

    print("AO3 EPUB inbox organizer (duplicates + interactive sorting)")
    print(f"Inbox: {ROOT}")
    print(f"Output: {OUT_ROOT}")
    print(f"Memory: {MEMORY_PATH.name}")
    print(f"Progress: {PROGRESS_PATH.name}")
    print(f"Run log: {RUN_LOG_PATH.name}")
    print()

    ensure_folder(OUT_ROOT)

    mem = load_memory()
    prog = load_progress()
    done = prog.get("done", {})

    updated_folders: Set[str] = set()

    existing_idx = build_existing_index()

    items = iter_inbox_items()
    if not items:
        print("No items found in inbox.")
        append_jsonl(
            RUN_LOG_PATH,
            {
                "run_id": run_id,
                "when": run_when,
                "action": "run_summary",
                "updated_folders": [],
                "note": "No inbox items",
            },
        )
        return

    for p in items:
        try:
            k = item_key(p)
            if k in done:
                continue

            info = extract_item_info(p)

            print("\n" + "=" * 80)
            print(f"Processing: {info.path.name}")

            # ------------------------------------------------------------
            # Step A: Route duplicates into existing main/ location (files only)
            # ------------------------------------------------------------
            if info.path.is_file():
                existing_file = find_existing_destination(
                    existing_idx, info.path.name, info.author
                )
                if existing_file is not None:
                    dest_folder = existing_file.parent

                    if existing_is_older(existing_file, info.itemdate):
                        remove_path(existing_file)
                        shutil.move(str(info.path), str(existing_file))
                        log_action(
                            run_id,
                            "routed_duplicate_replaced",
                            info.path.name,
                            str(dest_folder),
                            f"Replaced with newer at {existing_file.name}",
                        )
                        print(f"[Route+Replace] -> {existing_file.relative_to(ROOT)}")
                    else:
                        remove_path(info.path)
                        log_action(
                            run_id,
                            "routed_duplicate_discarded",
                            info.path.name,
                            str(dest_folder),
                            f"Existing newer; discarded incoming (kept {existing_file.name})",
                        )
                        print(
                            f"[Route+Discard] Existing newer; removed incoming: {info.path.name}"
                        )

                    updated_folders.add(str(dest_folder))
                    done[k] = {
                        "when": datetime.now().isoformat(timespec="seconds"),
                        "action": "routed_duplicate",
                        "dest_folder": str(dest_folder),
                    }
                    prog["done"] = done
                    save_progress(prog)

                    existing_idx = build_existing_index()
                    continue

            # ------------------------------------------------------------
            # Step B: Interactive sorting
            # ------------------------------------------------------------
            fandom = find_matching_fandom(info.fandomraw, mem)
            if fandom is None:
                fandom = prompt_for_fandom(info.path.name, info.fandomraw, mem)
                if fandom == "__SKIP__":
                    done[k] = {
                        "when": datetime.now().isoformat(timespec="seconds"),
                        "action": "skipped",
                    }
                    prog["done"] = done
                    save_progress(prog)
                    log_action(
                        run_id,
                        "skipped",
                        info.path.name,
                        str(ROOT),
                        "User chose skip",
                    )
                    continue
                if fandom == "__UNDECIDED__":
                    log_action(
                        run_id,
                        "undecided_fandom",
                        info.path.name,
                        str(ROOT),
                        "Left in inbox (user undecided)",
                    )
                    print("[Undecided fandom] Left in inbox.")
                    continue

                update_fandom_mem(mem, fandom, info.fandomraw)
                save_memory(mem)
                log_action(
                    run_id,
                    "memory_updated_fandom",
                    info.path.name,
                    str(ROOT),
                    f"Fandom set to {fandom}",
                )
                print(f"[Memory] Updated fandom mapping -> {fandom}")

            fandom_folder = ensure_folder(OUT_ROOT / sanitize_name(fandom))

            rel_folder = resolve_relationship(
                info.path.name, fandom, info.slashpairs, mem, info.relraw, info.summary
            )
            if rel_folder == "__SKIP__":
                done[k] = {
                    "when": datetime.now().isoformat(timespec="seconds"),
                    "action": "skipped",
                }
                prog["done"] = done
                save_progress(prog)
                log_action(
                    run_id,
                    "skipped",
                    info.path.name,
                    str(ROOT),
                    "User chose skip at relationship step",
                )
                continue

            dest_folder = fandom_folder
            if rel_folder is not None:
                rel_folder_clean = sanitize_name(rel_folder)
                remember_relationship(mem, fandom, rel_folder_clean)
                save_memory(mem)
                dest_folder = ensure_folder(fandom_folder / rel_folder_clean)
                log_action(
                    run_id,
                    "relationship_chosen",
                    info.path.name,
                    str(dest_folder),
                    f"Relationship folder = {rel_folder_clean}",
                )
            else:
                rel_folder_clean = None

            sub = choose_series_or_author_subfolder(info.text, info.author)
            if sub:
                dest_folder = ensure_folder(dest_folder / sub)

            # ------------------------------------------------------------
            # Move + conflict resolution
            # ------------------------------------------------------------
            target = dest_folder / info.path.name
            if not target.exists():
                shutil.move(str(info.path), str(target))
                updated_folders.add(str(dest_folder))
                log_action(
                    run_id,
                    "moved",
                    info.path.name,
                    str(dest_folder),
                    "No conflict",
                )
                print(f"[Move] -> {target.relative_to(ROOT)}")

                done[k] = {
                    "when": datetime.now().isoformat(timespec="seconds"),
                    "action": "moved",
                    "dest_folder": str(dest_folder),
                }
                prog["done"] = done
                save_progress(prog)

                existing_idx.setdefault(target.name.lower(), []).append(target)

                # --- NEW: remember fic in memory ---
                mem_fics = mem.setdefault("fics", [])
                mem_fics.append(
                    {
                        "local_path": str(target),
                        "title": info.title,
                        "ao3_url": info.ao3_url,
                        "fandom": fandom,
                        "relationship": rel_folder_clean,
                        "author": info.author,
                        "chapters_local": info.chapters_local,
                        "complete": info.complete,
                        "last_updated": datetime.now().isoformat(timespec="seconds"),
                    }
                )
                save_memory(mem)

                continue

            by_name = by_author_name(info.path.name, info.author)
            target2 = dest_folder / by_name

            if not target2.exists():
                shutil.move(str(info.path), str(target2))
                updated_folders.add(str(dest_folder))
                log_action(
                    run_id,
                    "moved_renamed_by_author",
                    info.path.name,
                    str(dest_folder),
                    f"Renamed to {by_name}",
                )
                print(f"[Move conflict] -> {target2.relative_to(ROOT)}")

                done[k] = {
                    "when": datetime.now().isoformat(timespec="seconds"),
                    "action": "moved_by_author",
                    "dest_folder": str(dest_folder),
                }
                prog["done"] = done
                save_progress(prog)

                existing_idx.setdefault(target2.name.lower(), []).append(target2)

                # --- NEW: remember fic in memory (renamed) ---
                mem_fics = mem.setdefault("fics", [])
                mem_fics.append(
                    {
                        "local_path": str(target2),
                        "title": info.title,
                        "ao3_url": info.ao3_url,
                        "fandom": fandom,
                        "relationship": rel_folder_clean,
                        "author": info.author,
                        "chapters_local": info.chapters_local,
                        "complete": info.complete,
                        "last_updated": datetime.now().isoformat(timespec="seconds"),
                    }
                )
                save_memory(mem)

                continue

            if existing_is_older(target2, info.itemdate):
                remove_path(target2)
                shutil.move(str(info.path), str(target2))
                updated_folders.add(str(dest_folder))
                log_action(
                    run_id,
                    "replaced_old_with_newer",
                    info.path.name,
                    str(dest_folder),
                    f"Kept newer at {target2.name}",
                )
                print(f"[Replace] Kept newer -> {target2.relative_to(ROOT)}")
                # Optional: update memory here too, similar to above
            else:
                remove_path(info.path)
                log_action(
                    run_id,
                    "discarded_incoming_older",
                    info.path.name,
                    str(dest_folder),
                    f"Existing newer at {target2.name}",
                )
                print(f"[Discard] Existing newer; removed incoming: {info.path.name}")

            done[k] = {
                "when": datetime.now().isoformat(timespec="seconds"),
                "action": "resolved_conflict",
                "dest_folder": str(dest_folder),
            }
            prog["done"] = done
            save_progress(prog)

            existing_idx = build_existing_index()

        except KeyboardInterrupt:
            print("\n[Stop] Interrupted. Progress saved. Re-run to continue.")
            save_memory(mem)
            prog["done"] = done
            save_progress(prog)
            append_jsonl(
                RUN_LOG_PATH,
                {
                    "run_id": run_id,
                    "when": datetime.now().isoformat(timespec="seconds"),
                    "action": "run_summary",
                    "updated_folders": sorted(updated_folders),
                    "note": "Interrupted by user",
                },
            )
            return
        except Exception as e:
            log_action(run_id, "error", p.name, str(ROOT), f"{type(e).__name__}: {e}")
            print(f"[Error] {p.name}: {e}")
            # not marking done so you can retry next run

    append_jsonl(
        RUN_LOG_PATH,
        {
            "run_id": run_id,
            "when": run_when,
            "action": "run_summary",
            "updated_folders": sorted(updated_folders),
            "note": "Completed",
        },
    )

    print("\nDone. Updated folders this run:")
    for f in sorted(updated_folders):
        print("  -", f)


if __name__ == "__main__":
    main()
