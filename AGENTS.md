# Guide for AI assistants

This file is for AI coding assistants (Claude Code, Codex, Cursor and others) helping a user
rearrange their iPhone home screen with this tool. The user gives instructions in plain English,
and you turn them into workbook edits, dry runs and pushes.

## The tool

- `iphone_organizer.py` is the GUI and CLI. Its `App` class holds the operations:
  `_read_workbook`, `_backup`, `_push` and `_restore`.
- `layout.py` does the conversion. `flatten(state)` turns the phone's icon state into rows,
  `build(rows, state)` returns `(new_state, hidden, warnings, errors)`, and
  `summarize_changes(before, after)` describes the difference.
- Output (workbooks, backups, logs, icons) lives under the **Output folder** set in
  `iphone_organizer.ini` (`output_dir`). If that is blank, it goes next to the script.
- Headless: `python iphone_organizer.py --headless {export,preview,push,restore}`.

## Workflow for every change

1. **Start from the newest workbook**, which should match the last push. If you're unsure,
   export a fresh one first (`--headless export`, or the Export button).
2. **Never edit a workbook in place.** Load it with openpyxl, change only the Target Page,
   Target Folder and Target Order columns, and save a **new** file named
   `YYYYMMDD-HHMMSS-iphone_apps-<short-change-name>.xlsx`.
3. **Find apps by bundle ID**, not by name alone. Names repeat (for example two apps called
   "Empresas"). The App Store category in the Category column helps to find groups such as
   games, airlines or food delivery.
4. **Ask before guessing.** When "my games" or "hotel apps" includes borderline apps (companion
   apps, booking sites), list them and let the user choose. Also ask before taking an app out of
   a folder it's already in.
5. **Dry run.** Read the live layout and build the new one without writing anything:
   `rows = app._read_workbook(path)`, get the live state, then run `layout.build(rows, state)`
   and `layout.summarize_changes(state, new)`. Report the page and folder results, any overflow,
   and anything unexpected. `--headless preview` does the same and prints the pages.
6. **Push only when the user explicitly says so.** Back up first (`app._backup(before,
   "before-push")`), then call `app._push(new_state, hidden)`, which writes, re-reads and
   verifies. Give the user the before-push backup path as the undo point.

## Placing things

- Target Order sets the position on a page, smallest first. Use decimals to slot something in
  (3.5 goes between 3 and 4).
- Every row with the same Target Folder name ends up in one folder. The folder sits at the
  position of its lowest-ordered member, and members are ordered by their own Target Order
  values. For a new folder after an existing one at 1.5, use 1.6 + n/1000 for its members.
- Renaming a folder means changing Target Folder on all of its rows.

## iOS limits to respect

- A page holds 24 icons, a folder page 9, and the dock 4. When a page has too many icons, the
  extras move to the start of the next page and can ripple through several full pages. Prefer
  taking icons from pages that have room, and tell the user what shifts.
- **15 pages maximum.** Anything past page 15 goes to the App Library only.
- **Leaving an app out doesn't hide it.** iOS puts every installed app missing from a pushed
  layout back onto a page. To take an app off the home screen, delete it or place it past page 15.
- **Deleted apps:** if the dry run says an app is being "added from the App Library" and you
  didn't intend that, check whether it's still installed. If the user deleted it, clear that
  row's Target columns in a new workbook copy.
- Widgets can't go into folders or the dock.

## User rules

Users often give standing rules ("page 1 is off limits", "new folders go right after the last
one I made"). Save them where you keep project instructions (for example a local `CLAUDE.md`)
and follow them on every later change.
