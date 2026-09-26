# iPhone Screen Organize

Rearrange your iPhone home screen from a Mac over USB, using an Excel workbook. No jailbreak, no MDM,
no device supervision.

Export your current layout to a spreadsheet, move apps between pages and folders by editing three
columns, preview the result with real app icons, then push it to the phone. Every push saves a
backup first, and one click restores it.

Tested on an iPhone 15 Pro Max running iOS 27.

## How it works

The phone runs a SpringBoard service that reads and writes the home screen layout. It's the
service iTunes once used for arranging apps. [pymobiledevice3](https://github.com/doronz88/pymobiledevice3)
talks to it over USB. This tool adds:

- an **Excel round trip**. Every app is one row, with its current page, folder and position, its
  App Store category, and a suggested folder based on that category.
- a **preview window** that draws every page with the app icons from your phone and highlights
  what would change.
- **backup, verify and restore** around every write.

## Requirements

- macOS (tested). pymobiledevice3 also supports Linux and Windows, but this tool hasn't been tried there.
- Python 3.11 or later
- An iPhone connected by USB, unlocked, with "Trust This Computer" accepted

```
python3 -m venv ~/venvs/iPhone
source ~/venvs/iPhone/bin/activate
pip install -r requirements.txt
```

## Usage

```
python iphone_organizer.py
```

1. **Export to Excel** reads the phone's layout and installed apps, looks up App Store categories,
   caches icons for the preview, and writes `excel/<timestamp>-iphone_apps.xlsx`.
2. Edit the yellow **Target Page**, **Target Folder** and **Target Order** columns. The workbook's
   "How To" sheet explains them. Save and close Excel.
3. **Preview** shows the resulting pages without touching the phone.
4. **Push to iPhone** backs up the current layout, shows the preview, writes on confirm, then
   reads the layout back and reports anything iOS changed.
5. **Restore Backup** writes any saved layout back to the phone.

Every app with the same Target Folder name ends up in one folder. Decimal Target Order values slot
an app between others (3.5 goes between 3 and 4).

### Let an AI assistant do the work

You don't have to edit the workbook yourself. Once you've exported it, an AI coding assistant
such as [Claude Code](https://claude.com/claude-code) can take instructions in plain English and
handle the rest. It finds the right apps, fills in the Target columns, checks the result and
pushes it when you say so. The **Use AI** button in the window shows the steps and a ready-to-paste first message with your
file paths filled in. Open the assistant in this project's folder and give it orders like:

- "Make a Travel folder first on page 2 with my airline and cruise apps."
- "Put all my Microsoft apps in one folder, but leave Edge where it is."
- "Find my food delivery apps and put them in a folder called Gig after Portuguese."
- "Move Authy to page 3 without shifting everything else."
- "Show me what would change before you push anything."

The assistant reads [AGENTS.md](AGENTS.md), which explains how to work safely with this tool:
edit a new copy of the workbook, do a dry run, and push only when you approve. You stay in
charge. It asks when an instruction is ambiguous, for example which borderline apps count as
"games". Every push is backed up, so Restore undoes it.

You can also give it standing rules to remember, like "page 1 is off limits unless I say so" or
"new folders go right after the last one I created".

### Without the window

```
python iphone_organizer.py --headless export [--no-icons]
python iphone_organizer.py --headless preview [--workbook PATH]
python iphone_organizer.py --headless push [--workbook PATH] [--yes]
python iphone_organizer.py --headless restore [--backup PATH] [--yes]
```

The preview is printed as text, with `*` marking apps that change page or folder. `push` and
`restore` ask for confirmation unless you pass `--yes`.

### Quick test first

`layout_probe.py` is a small standalone check. It reads your layout, swaps two apps on the last
page, reads it back and tells you whether iOS accepted the write. Restore puts it back. Run it
once before organizing, since Apple could restrict this service in a future iOS release.

## iOS behaviour to know about

- **15 pages maximum.** Anything laid out past page 15 goes to the App Library only. The preview
  and a warning flag it.
- **Apps can't be hidden by leaving them out.** iOS puts every installed app that is missing from a
  pushed layout back onto a page. To remove an app from the home screen, delete it or place it past
  page 15.
- A page holds 24 icons, a folder page holds 9, and the dock holds 4. When a page has too many,
  the extras move to the start of the next page.
- Widgets can move between pages but can't go into folders or the dock.

## Output folders

The tool creates these in the **Output folder** set at the top of the window, which is saved as
`output_dir` in `iphone_organizer.ini`. If it's blank, they go next to the script. They describe
your phone, so they are git-ignored:

| Folder | Contents |
| --- | --- |
| `backups/` | Restorable layout snapshots (`.plist`) |
| `excel/` | Exported workbooks |
| `json/` | Readable layout copies and cached App Store categories |
| `icons/` | App icons pulled from the phone for the preview |
| `logs/` | A timestamped log of every run |

## License

GPL-3.0, the same license as [pymobiledevice3](https://github.com/doronz88/pymobiledevice3), which this tool is built on. See [LICENSE](LICENSE). You may use, change and share it. If you distribute it or a modified version, you must publish the full source under the GPL too.
