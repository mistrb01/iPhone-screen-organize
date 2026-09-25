"""Probe whether the iPhone home screen layout can be read and written over USB.

Read Layout   - backs up the current layout and summarizes it.
Test Write    - swaps the first two apps on the last home screen page, then re-reads to verify.
Restore       - pushes a saved backup back to the phone.
"""

import asyncio
import configparser
import json
import plistlib
import re
import signal
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from pymobiledevice3.lockdown import create_using_usbmux
from pymobiledevice3.services.springboard import SpringBoardServicesService

HERE = Path(__file__).resolve().parent
LOG_DIR = HERE / "logs"
JSON_DIR = HERE / "json"
BACKUP_DIR = HERE / "backups"
INI_PATH = HERE / "layout_probe.ini"
SECTION = "layout_probe"
PATH_RE = re.compile(r"(/[^\s'\"]+\.(?:xlsx|csv|txt|json|plist|docx|zip))")


def ts() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def item_id(item) -> str | None:
    if isinstance(item, dict):
        return item.get("bundleIdentifier") or item.get("displayIdentifier")
    return None


def is_app(item) -> bool:
    return isinstance(item, dict) and "bundleIdentifier" in item and item.get("listType") != "folder"


def summarize(state: list) -> list[str]:
    lines = []
    for i, page in enumerate(state):
        label = "Dock" if i == 0 else f"Page {i}"
        apps = sum(1 for it in page if is_app(it))
        folders = [it for it in page if isinstance(it, dict) and it.get("listType") == "folder"]
        other = len(page) - apps - len(folders)
        lines.append(f"{label}: {len(page)} items ({apps} apps, {len(folders)} folders, {other} other/widgets)")
        for f in folders:
            n = sum(len(p) for p in f.get("iconLists", []))
            lines.append(f"    folder '{f.get('displayName')}' with {n} items")
    return lines


def to_jsonable(obj):
    if isinstance(obj, bytes):
        return f"<{len(obj)} bytes>"
    if isinstance(obj, datetime):
        return obj.isoformat()
    return str(obj)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.stop_flag = threading.Event()
        self.busy = False
        LOG_DIR.mkdir(exist_ok=True)
        self.log_path = LOG_DIR / f"{ts()}-layout_probe.txt"
        self.log_file = open(self.log_path, "a", encoding="utf-8")

        root.title("iPhone Layout Probe")
        root.geometry("900x600")

        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Backup to restore:").pack(side="left")
        self.backup_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.backup_var, width=80).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(top, text="Browse", command=self._browse).pack(side="left")

        self.log_w = scrolledtext.ScrolledText(root, wrap="word", font=("Menlo", 12))
        self.log_w.pack(fill="both", expand=True, padx=8)
        self.log_w.tag_config("link", foreground="blue", underline=True)
        self.log_w.tag_bind("link", "<Button-1>", self._open_link)
        self.log_w.tag_bind("link", "<Enter>", lambda e: self.log_w.config(cursor="hand2"))
        self.log_w.tag_bind("link", "<Leave>", lambda e: self.log_w.config(cursor=""))

        bar = ttk.Frame(root, padding=8)
        bar.pack(fill="x")
        for text, cmd in [
            ("Read Layout", lambda: self._run(self._read)),
            ("Test Write", lambda: self._run(self._test_write)),
            ("Restore Backup", lambda: self._run(self._restore)),
            ("Stop", self._stop),
            ("Save Config", self._save_config),
            ("Load Config", self._load_config),
            ("Copy Log", self._copy_log),
            ("Clear Log", lambda: self.log_w.delete("1.0", "end")),
        ]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)

        self._load_config(quiet=True)
        self.log(f"Log file: {self.log_path}")
        self.log("Plug in the iPhone, unlock it, and tap Trust if asked. Then click Read Layout.")

    # ---------- logging ----------
    def log(self, msg: str):
        line = f"{datetime.now().strftime('%Y%m%d-%H%M%S')}  {msg}"
        self.log_file.write(line + "\n")
        self.log_file.flush()
        self.root.after(0, self._append, line)

    def _append(self, line: str):
        start = self.log_w.index("end-1c")
        self.log_w.insert("end", line + "\n")
        for m in PATH_RE.finditer(line):
            if Path(m.group(1)).exists():
                self.log_w.tag_add("link", f"{start}+{m.start()}c", f"{start}+{m.end()}c")
        self.log_w.see("end")

    def _open_link(self, event):
        idx = self.log_w.index(f"@{event.x},{event.y}")
        rng = self.log_w.tag_prevrange("link", f"{idx}+1c")
        if not rng:
            return
        path = self.log_w.get(*rng)
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", path])
        elif sys.platform == "win32":
            subprocess.run(["explorer", f"/select,{path}"])

    def _copy_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log_w.get("1.0", "end-1c"))

    # ---------- config ----------
    def _save_config(self):
        cfg = configparser.ConfigParser()
        cfg.read(INI_PATH)
        cfg[SECTION] = {"backup_path": self.backup_var.get()}
        with open(INI_PATH, "w") as f:
            cfg.write(f)
        self.log(f"Config saved: {INI_PATH}")

    def _load_config(self, quiet=False):
        cfg = configparser.ConfigParser()
        cfg.read(INI_PATH)
        if cfg.has_section(SECTION):
            self.backup_var.set(cfg[SECTION].get("backup_path", ""))
        if not quiet:
            self.log("Config loaded.")

    def _browse(self):
        p = filedialog.askopenfilename(initialdir=BACKUP_DIR, filetypes=[("plist", "*.plist")])
        if p:
            self.backup_var.set(p)

    # ---------- run plumbing ----------
    def _stop(self):
        self.stop_flag.set()
        self.log("Stop requested. The current step will finish, then the operation ends.")

    def _run(self, coro_fn):
        if self.busy:
            self.log("Already running.")
            return
        self.busy = True
        self.stop_flag.clear()

        def worker():
            try:
                asyncio.run(coro_fn())
            except Exception as e:
                self.log(f"ERROR: {type(e).__name__}: {e}")
            finally:
                self.busy = False

        threading.Thread(target=worker, daemon=True).start()

    def _ask(self, title: str, msg: str) -> bool:
        done, answer = threading.Event(), []

        def show():
            answer.append(messagebox.askokcancel(title, msg))
            done.set()

        self.root.after(0, show)
        done.wait()
        return answer[0]

    async def _connect(self):
        self.log("Connecting over USB...")
        lockdown = await create_using_usbmux()
        info = lockdown.short_info
        self.log(f"Connected: {info.get('DeviceName')} ({info.get('ProductType')}), iOS {info.get('ProductVersion')}")
        return lockdown, SpringBoardServicesService(lockdown)

    def _backup(self, state: list, tag: str) -> Path:
        BACKUP_DIR.mkdir(exist_ok=True)
        JSON_DIR.mkdir(exist_ok=True)
        stamp = ts()
        plist_path = BACKUP_DIR / f"{stamp}-icon_state-{tag}.plist"
        plist_path.write_bytes(plistlib.dumps(state, fmt=plistlib.FMT_BINARY))
        json_path = JSON_DIR / f"{stamp}-icon_state-{tag}.json"
        json_path.write_text(json.dumps(state, indent=2, default=to_jsonable), encoding="utf-8")
        self.log(f"Backup (restorable): {plist_path}")
        self.log(f"Readable copy: {json_path}")
        return plist_path

    # ---------- operations ----------
    async def _read(self):
        lockdown, sb = await self._connect()
        try:
            state = await sb.get_icon_state()
            path = self._backup(state, "read")
            self.root.after(0, self.backup_var.set, str(path))
            for line in summarize(state):
                self.log(line)
        finally:
            await lockdown.close()

    async def _test_write(self):
        lockdown, sb = await self._connect()
        try:
            state = await sb.get_icon_state()
            path = self._backup(state, "before-test")
            self.root.after(0, self.backup_var.set, str(path))

            target = None
            for pi in range(len(state) - 1, 0, -1):
                idxs = [i for i, it in enumerate(state[pi]) if is_app(it)]
                if len(idxs) >= 2:
                    target = (pi, idxs[0], idxs[1])
                    break
            if not target:
                self.log("No home screen page has two plain apps to swap. Nothing written.")
                return
            pi, a, b = target
            name_a = state[pi][a].get("displayName")
            name_b = state[pi][b].get("displayName")
            if not self._ask(
                "Test Write",
                f"Swap '{name_a}' and '{name_b}' on page {pi}?\n\nBackup saved to:\n{path}",
            ):
                self.log("Test write cancelled.")
                return
            if self.stop_flag.is_set():
                return

            new_state = [list(p) for p in state]
            new_state[pi][a], new_state[pi][b] = new_state[pi][b], new_state[pi][a]
            self.log(f"Writing layout with '{name_a}' and '{name_b}' swapped on page {pi}...")
            await sb.set_icon_state(new_state)

            await asyncio.sleep(2)
            after = await sb.get_icon_state()
            self._backup(after, "after-test")
            want = [item_id(it) for it in new_state[pi]]
            got = [item_id(it) for it in after[pi]] if pi < len(after) else []
            before = [item_id(it) for it in state[pi]]
            if got == want:
                self.log("RESULT: write APPLIED. The phone accepted the new layout. Check the phone to confirm.")
            elif got == before:
                self.log("RESULT: write IGNORED. The layout is unchanged.")
            else:
                self.log("RESULT: write PARTIAL or reshuffled. Compare the before/after JSON files.")
            self.log("Use Restore Backup to put the original layout back.")
        finally:
            await lockdown.close()

    async def _restore(self):
        path = Path(self.backup_var.get())
        if not path.exists():
            self.log(f"Backup not found: {path}")
            return
        state = plistlib.loads(path.read_bytes())
        lockdown, sb = await self._connect()
        try:
            self.log(f"Restoring {path} ...")
            await sb.set_icon_state(state)
            await asyncio.sleep(2)
            after = await sb.get_icon_state()
            ok = [[item_id(i) for i in p] for p in after] == [[item_id(i) for i in p] for p in state]
            self.log("Restore verified." if ok else "Restore sent, but the re-read layout differs. Check the phone.")
        finally:
            await lockdown.close()

    def close(self):
        self.stop_flag.set()
        self.log_file.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = App(root)
    presses = {"n": 0}

    def on_sigint(*_):
        presses["n"] += 1
        if presses["n"] >= 3:
            sys.exit(1)
        app._stop()
        if presses["n"] == 2 or not app.busy:
            root.after(0, app.close)

    signal.signal(signal.SIGINT, on_sigint)
    root.protocol("WM_DELETE_WINDOW", app.close)
    def tick():  # lets Python see Ctrl-C while Tk is idle
        root.after(250, tick)

    tick()
    root.mainloop()


if __name__ == "__main__":
    main()
