"""Convert between the SpringBoard icon state and flat spreadsheet rows.

Icon state: list of pages, page 0 is the dock. Each item is an app dict, a widget dict
(has "iconType"), or a folder dict ({"displayName", "listType": "folder", "iconLists": [[...], ...]}).

Row fields: key, name, kind (App/Widget), page ("Dock", int, or None = App Library),
folder (str or None), order (float or None). Folder members get order = folder position
+ index/1000, so an unedited sheet rebuilds the same layout.
"""

from collections import defaultdict

PAGE_SLOTS = 24
DOCK_SLOTS = 4
FOLDER_PAGE_SLOTS = 9
MAX_PAGES = 15  # iOS puts anything past this into the App Library only


def item_key(item: dict) -> str | None:
    return item.get("displayIdentifier") or item.get("bundleIdentifier")


def is_folder(item: dict) -> bool:
    return item.get("listType") == "folder"


def kind_of(item: dict) -> str:
    return "Widget" if "iconType" in item else "App"


def page_label(i: int):
    return "Dock" if i == 0 else i


def flatten(state: list) -> list[dict]:
    rows = []
    for pi, page in enumerate(state):
        for pos, item in enumerate(page, start=1):
            if is_folder(item):
                members = [m for sub in item.get("iconLists", []) for m in sub]
                for mi, m in enumerate(members, start=1):
                    rows.append(_row(m, pi, item.get("displayName"), pos + mi / 1000))
            else:
                rows.append(_row(item, pi, None, pos))
    return rows


def _row(item: dict, pi: int, folder: str | None, order: float) -> dict:
    return {
        "key": item_key(item),
        "bundle_id": item.get("bundleIdentifier") or item.get("displayIdentifier"),
        "name": item.get("displayName") or f"Widget ({item.get('bundleIdentifier')})",
        "kind": kind_of(item),
        "page": page_label(pi),
        "folder": folder,
        "order": order,
    }


def items_by_key(state: list) -> dict[str, dict]:
    out = {}
    for page in state:
        for item in page:
            if is_folder(item):
                for sub in item.get("iconLists", []):
                    for m in sub:
                        out[item_key(m)] = m
            else:
                out[item_key(item)] = item
    return out


def _page_sort(p) -> int:
    return 0 if p == "Dock" else int(p)


def build(rows: list[dict], current_state: list) -> tuple[list, list[str], list[str], list[str]]:
    """Build a new icon state from target rows.

    Returns (state, hidden_names, warnings, errors). Items on the phone that are missing from
    the rows are appended to the last page, so nothing disappears by accident.
    """
    items = items_by_key(current_state)
    warnings, errors, hidden = [], [], []
    pages: dict = defaultdict(list)
    folders: dict = {}
    seen = set()

    for n, r in enumerate(rows):
        key = r["key"]
        if key in seen:
            errors.append(f"'{r['name']}' ({key}) is listed twice.")
            continue
        seen.add(key)
        item = items.get(key) or {
            "bundleIdentifier": r["bundle_id"],
            "displayIdentifier": r["bundle_id"],
            "displayName": r["name"],
        }
        if r["page"] is None:
            if r["kind"] == "Widget":
                warnings.append(f"Widget '{r['name']}' has no Target Page, so it will be removed.")
            hidden.append(r["name"])
            continue
        tb = (r["order"] if r["order"] is not None else 999, n)
        if r["folder"] and r["kind"] == "Widget":
            warnings.append(f"Widget '{r['name']}' can't go in a folder. Placed loose on its page.")
        if r["folder"] and r["kind"] != "Widget":
            f = folders.setdefault(r["folder"], {"members": [], "where": []})
            f["members"].append((tb, item))
            f["where"].append((_page_sort(r["page"]), tb))
        else:
            pages[_page_sort(r["page"])].append((tb, item))

    missing = [it for k, it in items.items() if k not in seen]
    if missing:
        last = max([p for p in pages if p > 0], default=1)
        for i, it in enumerate(missing):
            pages[last].append(((10_000 + i, 0), it))
        warnings.append(
            f"{len(missing)} item(s) on the phone are not in the workbook and were added to page {last}: "
            + ", ".join(it.get("displayName", "?") for it in missing)
        )

    for name, f in folders.items():
        page, tb = min(f["where"])
        if len({w[0] for w in f["where"]}) > 1:
            warnings.append(f"Folder '{name}' has members on several Target Pages. It goes on page {page or 'Dock'}.")
        members = [it for _, it in sorted(f["members"], key=lambda x: x[0])]
        chunks = [members[i:i + FOLDER_PAGE_SLOTS] for i in range(0, len(members), FOLDER_PAGE_SLOTS)]
        pages[page].append((tb, {"displayName": name, "listType": "folder", "iconLists": chunks}))

    dock = [it for _, it in sorted(pages.pop(0, []), key=lambda x: x[0])]
    if len(dock) > DOCK_SLOTS:
        errors.append(f"The dock has {len(dock)} items. The limit is {DOCK_SLOTS}.")
    if any(kind_of(it) == "Widget" for it in dock):
        errors.append("Widgets can't go in the dock.")

    state = [dock]
    carry = []
    for p in sorted(pages):
        ordered = carry + [it for _, it in sorted(pages[p], key=lambda x: x[0])]
        if len(ordered) > PAGE_SLOTS:
            warnings.append(f"Page {p} has {len(ordered)} items. The last {len(ordered) - PAGE_SLOTS} "
                            "move to the start of the next page.")
        state.append(ordered[:PAGE_SLOTS])
        carry = ordered[PAGE_SLOTS:]
    while carry:
        state.append(carry[:PAGE_SLOTS])
        carry = carry[PAGE_SLOTS:]
    if len(state) - 1 > MAX_PAGES:
        extra = [it.get("displayName", "?") for pg in state[MAX_PAGES + 1:] for it in pg]
        warnings.append(f"iOS allows {MAX_PAGES} pages. These {len(extra)} icons on later pages will go to the "
                        "App Library only: " + ", ".join(extra))
    return state, hidden, warnings, errors


def signature(state: list) -> list:
    """Order-sensitive layout fingerprint for comparing two states."""
    out = []
    for page in state:
        sig = []
        for it in page:
            if is_folder(it):
                sig.append(("folder", it.get("displayName"), tuple(item_key(m) for sub in it["iconLists"] for m in sub)))
            else:
                sig.append(item_key(it))
        out.append(sig)
    return out


def summarize_changes(before: list, after: list) -> list[str]:
    old = {r["key"]: r for r in flatten(before)}
    new = {r["key"]: r for r in flatten(after)}
    moved = sum(1 for k, r in new.items() if k in old and (old[k]["page"], old[k]["folder"]) != (r["page"], r["folder"]))
    added = sum(1 for k in new if k not in old)
    old_f = [it["displayName"] for p in before for it in p if is_folder(it)]
    new_f = [it["displayName"] for p in after for it in p if is_folder(it)]
    lines = [
        f"Pages: {len(before) - 1} now, {len(after) - 1} after.",
        f"Folders: {len(old_f)} now, {len(new_f)} after.",
        f"Apps changing page or folder: {moved}.",
    ]
    if added:
        lines.append(f"Apps added to the home screen from the App Library: {added}.")
    removed_apps = [r["name"] for k, r in old.items() if k not in new]
    if removed_apps:
        lines.append(f"Apps leaving the home screen (still in the App Library): {len(removed_apps)}.")
    created = sorted(set(new_f) - set(old_f))
    removed = sorted(set(old_f) - set(new_f))
    dup = sorted({n for n in old_f if old_f.count(n) > 1} & set(new_f))
    if created:
        lines.append("New folders: " + ", ".join(created))
    if removed:
        lines.append("Folders removed: " + ", ".join(removed))
    if dup:
        lines.append("Same-name folders merged: " + ", ".join(dup))
    return lines
