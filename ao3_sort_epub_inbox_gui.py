import tkinter as tk
from tkinter import messagebox, simpledialog
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Set
import shutil

# ============================================================
# Import shared logic from the original script
# ============================================================
from ao3_sort_epub_inbox import (
    ROOT,
    OUT_ROOT,
    MEMORY_PATH,
    PROGRESS_PATH,
    RUN_LOG_PATH,
    REL_LOG_PATH,
    SUPPORTED_FILE_EXTS,
    load_memory,
    save_memory,
    load_progress,
    save_progress,
    log_action,
    extract_item_info,
    ensure_folder,
    iter_inbox_items,
    build_existing_index,
    find_existing_destination,
    existing_is_older,
    by_author_name,
    choose_series_or_author_subfolder,
    ItemInfo,  # dataclass with title, ao3_url, chapters_local, complete
)

# ============================================================
# State object and minimal window
# ============================================================
class SortState:
    def __init__(self):
        self.run_id = datetime.now().strftime("%Y%m%dT%H%M%S")
        self.run_when = datetime.now().isoformat(timespec="seconds")

        self.mem = load_memory()
        self.prog = load_progress()
        self.done = self.prog.get("done", {})

        self.updated_folders: Set[str] = set()
        self.existing_idx = build_existing_index()
        self.items: List[Path] = iter_inbox_items()
        self.current_index = 0
        self.current_info: Optional[ItemInfo] = None

        self.stage = "fandom"  # "fandom", "fandom_more_info", "relationship"
        self.history: List[dict] = []  # for Back
        self.chosen_fandom: Optional[str] = None
        self.chosen_relationship: Optional[str] = None

    def save_all(self):
        self.prog["done"] = self.done
        save_progress(self.prog)
        save_memory(self.mem)


state = SortState()

# ============================================================
# Basic GUI window and labels
# ============================================================
root = tk.Tk()
root.title("AO3 EPUB Inbox Organizer (GUI)")

lbl_item = tk.Label(root, text="", font=("Segoe UI", 10, "bold"))
lbl_fandom = tk.Label(root, text="", wraplength=600, justify="left")
lbl_rel = tk.Label(root, text="", wraplength=600, justify="left")
lbl_summary = tk.Label(root, text="", wraplength=600, justify="left")

lbl_item.grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))
lbl_fandom.grid(row=1, column=0, columnspan=4, sticky="w", padx=8, pady=2)
lbl_rel.grid(row=2, column=0, columnspan=4, sticky="w", padx=8, pady=2)
lbl_summary.grid(row=3, column=0, columnspan=4, sticky="w", padx=8, pady=(2, 8))

# ============================================================
# Fandom / relationship buttons (we’ll reuse them per stage)
# ============================================================
btn_1 = tk.Button(root, text="1) ", width=30)
btn_2 = tk.Button(root, text="2) ", width=30)
btn_3 = tk.Button(root, text="3) ", width=30)
btn_4 = tk.Button(root, text="4) ", width=30)
btn_5 = tk.Button(root, text="5) ", width=30)

btn_back = tk.Button(root, text="B) Back", width=20)
btn_save_quit = tk.Button(root, text="Q) Save progress & quit", width=25)

btn_1.grid(row=4, column=0, sticky="ew", padx=8, pady=2)
btn_2.grid(row=4, column=1, sticky="ew", padx=8, pady=2)
btn_3.grid(row=5, column=0, sticky="ew", padx=8, pady=2)
btn_4.grid(row=5, column=1, sticky="ew", padx=8, pady=2)
btn_5.grid(row=6, column=0, sticky="ew", padx=8, pady=2)

btn_back.grid(row=6, column=1, sticky="ew", padx=8, pady=2)
btn_save_quit.grid(row=7, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 8))

# define generic handlers
def go_back():
    if not state.history:
        return
    snap = state.history.pop()
    state.stage = snap["stage"]
    state.current_index = snap["current_index"]
    state.current_info = snap["current_info"]
    state.chosen_fandom = snap.get("chosen_fandom")
    state.chosen_relationship = snap.get("chosen_relationship")
    update_view()


def save_and_quit():
    state.save_all()
    root.destroy()


btn_back.config(command=go_back)
btn_save_quit.config(command=save_and_quit)

root.bind("b", lambda e: go_back())
root.bind("B", lambda e: go_back())
root.bind("q", lambda e: save_and_quit())
root.bind("Q", lambda e: save_and_quit())

# ============================================================
# Load the current item and show fandom tags
# ============================================================
def advance_to_next_item():
    # skip items marked done in progress
    while state.current_index < len(state.items):
        p = state.items[state.current_index]
        k = f"{'DIR' if p.is_dir() else 'FILE'}|{p.name}|{int(p.stat().st_mtime)}"
        if k in state.done:
            state.current_index += 1
            continue
        state.current_info = extract_item_info(p)
        state.stage = "fandom"
        state.chosen_fandom = None
        state.chosen_relationship = None
        state.history.clear()
        update_view()
        return

    # no more items
    messagebox.showinfo("Done", "No more inbox items to process.")
    state.save_all()
    root.destroy()

# ============================================================
# Implement fandom-stage handlers
# ============================================================
def snapshot_state():
    state.history.append(
        {
            "stage": state.stage,
            "current_index": state.current_index,
            "current_info": state.current_info,
            "chosen_fandom": state.chosen_fandom,
            "chosen_relationship": state.chosen_relationship,
        }
    )


def fandom_choose_existing():
    existing = sorted(state.mem.get("fandoms", {}).keys(), key=lambda x: x.lower())
    if not existing:
        messagebox.showinfo("No fandoms", "No fandoms in memory yet.")
        return

    options = "\n".join(f"{i+1}. {name}" for i, name in enumerate(existing))
    ans = simpledialog.askstring(
        "Existing fandom",
        f"Existing fandoms:\n{options}\n\nEnter number or name:",
        parent=root,
    )
    if not ans:
        return

    choice = None
    if ans.isdigit():
        idx = int(ans) - 1
        if 0 <= idx < len(existing):
            choice = existing[idx]
    else:
        for name in existing:
            if name.lower() == ans.strip().lower():
                choice = name
                break

    if not choice:
        messagebox.showerror("Invalid", "Invalid choice.")
        return

    snapshot_state()
    state.chosen_fandom = choice
    state.stage = "relationship"
    update_view()


def fandom_new():
    name = simpledialog.askstring("New fandom", "New fandom folder name:", parent=root)
    if not name:
        return
    name = name.strip()
    if not name:
        messagebox.showerror("Invalid", "Name cannot be empty.")
        return

    snapshot_state()
    state.chosen_fandom = name
    from ao3_sort_epub_inbox import update_fandom_mem

    update_fandom_mem(state.mem, name, state.current_info.fandomraw or "")
    save_memory(state.mem)
    state.stage = "relationship"
    update_view()


def fandom_undecided():
    info = state.current_info
    log_action(
        state.run_id,
        "undecided_fandom",
        info.path.name,
        str(ROOT),
        "Left in inbox (user undecided)",
    )
    messagebox.showinfo(
        "Undecided", "Item left in inbox; will be asked again next run."
    )
    state.current_index += 1
    advance_to_next_item()


def fandom_skip():
    info = state.current_info
    k = f"{'DIR' if info.path.is_dir() else 'FILE'}|{info.path.name}|{int(info.path.stat().st_mtime)}"
    state.done[k] = {
        "when": datetime.now().isoformat(timespec="seconds"),
        "action": "skipped",
    }
    log_action(
        state.run_id,
        "skipped",
        info.path.name,
        str(ROOT),
        "User chose skip",
    )
    state.current_index += 1
    advance_to_next_item()


def fandom_more_info():
    snapshot_state()
    state.stage = "fandom_more_info"
    update_view()


def fandom_back_basic():
    if state.history:
        state.history.pop()
    state.stage = "fandom"
    update_view()

# Each Update View
def update_view():
    info = state.current_info
    if info is None:
        lbl_item.config(text="No item loaded.")
        lbl_fandom.config(text="")
        lbl_rel.config(text="")
        lbl_summary.config(text="")
        return

    lbl_item.config(text=f"Item: {info.path.name}")

    if state.stage == "fandom":
        lbl_fandom.config(
            text=f"Fandom tag(s): {info.fandomraw or '[none detected]'}"
        )
        lbl_rel.config(text="")
        lbl_summary.config(text="")

        btn_1.config(text="1) Choose existing fandom", command=fandom_choose_existing)
        btn_2.config(text="2) New fandom", command=fandom_new)
        btn_3.config(
            text="3) Undecided (leave in inbox)", command=fandom_undecided
        )
        btn_4.config(text="4) Skip item", command=fandom_skip)
        btn_5.config(
            text="5) Require more information", command=fandom_more_info
        )

    elif state.stage == "fandom_more_info":
        lbl_fandom.config(
            text=f"Fandom tag(s): {info.fandomraw or '[none detected]'}"
        )
        lbl_rel.config(
            text=f"Relationship tag(s): {info.relraw or '[none detected]'}"
        )
        lbl_summary.config(
            text=f"Summary:\n{info.summary or '[no summary found]'}"
        )

        btn_1.config(text="1) Choose existing fandom", command=fandom_choose_existing)
        btn_2.config(text="2) New fandom", command=fandom_new)
        btn_3.config(
            text="3) Undecided (leave in inbox)", command=fandom_undecided
        )
        btn_4.config(text="4) Skip item", command=fandom_skip)
        btn_5.config(text="5) Back to basic view", command=fandom_back_basic)

    elif state.stage == "relationship":
        lbl_fandom.config(
            text=f"Chosen fandom folder: {state.chosen_fandom}"
        )
        lbl_rel.config(
            text=f"Relationship tag(s): {info.relraw or '[none detected]'}"
        )
        lbl_summary.config(
            text=f"Summary:\n{info.summary or '[no summary found]'}"
        )

        btn_1.config(
            text="1) Keep in fandom (no relationship)",
            command=rel_keep_in_fandom,
        )
        btn_2.config(
            text="2) Choose existing relationship folder",
            command=rel_choose_existing,
        )
        btn_3.config(
            text="3) Pick from slash-pairs",
            command=rel_pick_from_pairs,
        )
        btn_4.config(
            text="4) New relationship folder",
            command=rel_new,
        )
        btn_5.config(text="5) Skip item", command=rel_skip)


def click_btn_1(event=None):
    btn_1.invoke()


def click_btn_2(event=None):
    btn_2.invoke()


def click_btn_3(event=None):
    btn_3.invoke()


def click_btn_4(event=None):
    btn_4.invoke()


def click_btn_5(event=None):
    btn_5.invoke()


root.bind("1", click_btn_1)
root.bind("2", click_btn_2)
root.bind("3", click_btn_3)
root.bind("4", click_btn_4)
root.bind("5", click_btn_5)

# ============================================================
# Implement relationship-stage handlers and moving files
# ============================================================
def finalize_item_move(rel_folder: Optional[str]):
    info = state.current_info
    fandom = state.chosen_fandom
    if not fandom or info is None:
        messagebox.showerror("Error", "No fandom/item chosen.")
        return

    fandom_folder = ensure_folder(OUT_ROOT / fandom)
    dest_folder = fandom_folder

    rel_folder_clean: Optional[str] = None
    if rel_folder:
        rel_folder_clean = rel_folder
        from ao3_sort_epub_inbox import remember_relationship

        remember_relationship(state.mem, fandom, rel_folder_clean)
        save_memory(state.mem)
        dest_folder = ensure_folder(fandom_folder / rel_folder_clean)
        log_action(
            state.run_id,
            "relationship_chosen",
            info.path.name,
            str(dest_folder),
            f"Relationship folder = {rel_folder_clean}",
        )

    sub = choose_series_or_author_subfolder(info.text, info.author)
    if sub:
        dest_folder = ensure_folder(dest_folder / sub)

    target = dest_folder / info.path.name
    if not target.exists():
        shutil.move(str(info.path), str(target))
        state.updated_folders.add(str(dest_folder))
        log_action(
            state.run_id, "moved", info.path.name, str(dest_folder), "No conflict"
        )

        k = f"{'DIR' if info.path.is_dir() else 'FILE'}|{info.path.name}|{int(info.path.stat().st_mtime)}"
        state.done[k] = {
            "when": datetime.now().isoformat(timespec="seconds"),
            "action": "moved",
            "dest_folder": str(dest_folder),
        }
        state.existing_idx.setdefault(target.name.lower(), []).append(target)

        # remember fic in memory
        mem_fics = state.mem.setdefault("fics", [])
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
        save_memory(state.mem)

    else:
        from ao3_sort_epub_inbox import existing_is_older

        if existing_is_older(target, info.itemdate):
            target.unlink()
            shutil.move(str(info.path), str(target))
            log_action(
                state.run_id,
                "replaced_old_with_newer",
                info.path.name,
                str(dest_folder),
                f"Kept newer at {target.name}",
            )
            # Optionally update mem["fics"] here if you want to track this replacement
        else:
            info.path.unlink()
            log_action(
                state.run_id,
                "discarded_incoming_older",
                info.path.name,
                str(dest_folder),
                f"Existing newer at {target.name}",
            )

    state.current_index += 1
    advance_to_next_item()


def rel_keep_in_fandom():
    snapshot_state()
    finalize_item_move(None)


def rel_choose_existing():
    existing = state.mem.get("relationships_by_fandom", {}).get(
        state.chosen_fandom, []
    )
    if not existing:
        messagebox.showinfo(
            "No relationships",
            "No known relationship folders for this fandom.",
        )
        return
    options = "\n".join(f"{i+1}. {name}" for i, name in enumerate(existing))
    ans = simpledialog.askstring(
        "Existing relationship",
        f"Known relationship folders:\n{options}\n\nEnter number or name:",
        parent=root,
    )
    if not ans:
        return
    choice = None
    if ans.isdigit():
        idx = int(ans) - 1
        if 0 <= idx < len(existing):
            choice = existing[idx]
    else:
        for name in existing:
            if name.lower() == ans.strip().lower():
                choice = name
                break
    if not choice:
        messagebox.showerror("Invalid", "Invalid choice.")
        return
    snapshot_state()
    finalize_item_move(choice)


def rel_pick_from_pairs():
    pairs = state.current_info.slashpairs
    if not pairs:
        messagebox.showinfo(
            "No pairs", "No slash-pairs detected for this item."
        )
        return
    options = "\n".join(f"{i+1}. {p}" for i, p in enumerate(pairs))
    ans = simpledialog.askstring(
        "Slash-pair",
        f"Detected pairs:\n{options}\n\nEnter number:",
        parent=root,
    )
    if not ans or not ans.isdigit():
        return
    idx = int(ans) - 1
    if not (0 <= idx < len(pairs)):
        messagebox.showerror("Invalid", "Number out of range.")
        return
    snapshot_state()
    finalize_item_move(pairs[idx])


def rel_new():
    name = simpledialog.askstring(
        "New relationship", "New relationship folder name:", parent=root
    )
    if not name:
        return
    name = name.strip()
    if not name:
        messagebox.showerror("Invalid", "Name cannot be empty.")
        return
    snapshot_state()
    finalize_item_move(name)


def rel_skip():
    info = state.current_info
    k = f"{'DIR' if info.path.is_dir() else 'FILE'}|{info.path.name}|{int(info.path.stat().st_mtime)}"
    state.done[k] = {
        "when": datetime.now().isoformat(timespec="seconds"),
        "action": "skipped",
    }
    log_action(
        state.run_id,
        "skipped",
        info.path.name,
        str(ROOT),
        "User chose skip at relationship step",
    )
    state.current_index += 1
    advance_to_next_item()

# ============================================================
# Main
# ============================================================
def main():
    advance_to_next_item()
    root.mainloop()


if __name__ == "__main__":
    main()
