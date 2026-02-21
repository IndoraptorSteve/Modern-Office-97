```
 __  __ _                           __ _      ___   __  __ _            ___ _____ 
|  \/  (_)__ _ _ ___ ___ ___ / _| |_   / _ \ / _|/ _(_)__ ___   / _ \___  |
| |\/| | / _| '_/ _ (_-</ _ \  _|  _| | (_) |  _|  _| / _/ -_) | (_) | / / 
|_|  |_|_\__|_| \___/__/\___/_|  \__|  \___/|_| |_| |_\__\___|  \___/ /_/  
```

# Microsoft Office 97 — Electron Recreation

A faithful recreation of Microsoft Office 97 Professional Edition, built with Electron. Every detail — from the installer wizard to the toolbar icons — is designed to match the original 1997 release as closely as possible.

## Screenshots

*Coming soon*

## What's Included

### Setup Wizard
- Authentic multi-step installer matching the real Office 97 setup
- CD Key entry, Name & Organization, Installation Type (Typical/Custom/Run from CD-ROM)
- Install location with disk space info
- File copy progress with rotating feature ads on teal background
- Puzzle piece animation

### Microsoft Word 97
- Full document editor with contentEditable
- Interactive ruler with draggable margin/indent markers and tab stops
- Font & size dropdowns, Bold/Italic/Underline formatting
- Find & Replace across the document
- Page Setup dialog with margins, paper size, orientation
- Normal and Page Layout views
- Document Map, spell check, word count
- File open/save (.docx via mammoth/docx libraries)
- Print support

### Microsoft Excel 97
- 500 rows x 104 columns (A through CZ) spreadsheet grid
- Formula engine with cell references (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Multi-letter column support (AA, AB, ... CZ)
- Chart Wizard — 4-step wizard for bar, line, pie, scatter, area charts
- Sort dialog, AutoFill via drag handle
- Format Cells dialog (number formats, borders, alignment)
- Goal Seek, freeze panes, auto-filter
- File open/save (.xlsx via SheetJS)

### Microsoft PowerPoint 97
- Slide editor with 720x540 coordinate space, scalable zoom
- **242 design templates** extracted from original .pot files (real assets, not approximations)
- Template placeholder text boxes at exact original positions
- Click-and-drag shape drawing (Line, Arrow, Rectangle, Oval, Text Box)
- **AutoShapes** — 30+ shapes across 6 categories (Basic Shapes, Block Arrows, Flowchart, Stars & Banners, Callouts, Lines)
- **Fill Effects** — Gradients (24 PP97 presets), Textures (24), Patterns (48 hatch styles)
- **3-D Effects** — 20 preset styles with depth, direction, extrusion color, lighting
- **Shadow Effects** — 10 preset shadow styles
- Format AutoShape dialog with Colors/Lines, Size, Position, 3-D tabs
- Office 97 standard 40-color picker
- Chart insertion with full Chart Wizard (column, bar, line, pie, scatter, area, doughnut)
- Table insertion with cell editing
- **200 original Office 97 clipart** (SVG) via Clip Gallery 3.0 dialog
- Slide Sorter view with drag-and-drop reordering
- Notes pane, Outline view
- Slide transitions (15+ effects) with slideshow mode
- Find & Replace across all slides
- Print support (one slide per page)
- Undo/Redo with full state tracking
- Right-click context menus on slides and objects
- Keyboard shortcuts (Ctrl+C/X/V, arrow keys, Delete, F5, etc.)
- File open/save (.pptx via pptxgenjs)

### Microsoft Access 97
- Database browser with table, form, and query views
- Table datasheet view and design view
- Form view for data entry
- Query builder with SQL editor
- Import/Export CSV
- Sample database templates (Contacts, Inventory, etc.)

### Microsoft Outlook 97
- Inbox with sample emails
- Calendar with day and month views
- Contacts CRUD
- Tasks manager
- Compose, Reply, Forward
- Services setup wizard

## UI Authenticity

- **Win95/97 UI framework** — Custom `win95-ui.css` and `win95-ui.js` providing authentic Windows 95 controls (buttons, dialogs, toolbars, menus, scrollbars, tabs)
- **Original toolbar icons** — Extracted from real Office 97 binaries
- **Original app icons** — Word, Excel, PowerPoint, Access, Outlook
- **Pixelated font rendering** — LCD text and font hinting disabled for period-accurate look
- **Proper menu bar** — File, Edit, View, Insert, Format, Tools, Window, Help (varies per app)
- **Drawing toolbar** — Horizontal at bottom with AutoShapes, color pickers, shadow/3-D
- **Formatting toolbar** — Font/Size dropdowns, B/I/U, alignment, bullets/numbering
- **Status bar** — Slide count, view mode buttons, zoom

## Tech Stack

- **Electron 28** — Desktop app framework
- **Vanilla JavaScript** — No React, no frameworks, just like the 90s
- **HTML/CSS** — Custom Win95 UI components
- **mammoth / docx** — Word document import/export
- **SheetJS (xlsx)** — Excel file import/export
- **pptxgenjs** — PowerPoint file export

## Building

```bash
# Install dependencies
npm install

# Run in development
npm start

# Run a specific app
npm run word
npm run excel
npm run powerpoint

# Build Windows installer (.exe)
npm run build
```

The build outputs to `dist/Microsoft Office 97 Setup 1.0.0.exe`.

## Project Structure

```
office97/
├── main.js                    # Electron main process
├── launcher.html              # App launcher (Office Shortcut Bar)
├── installer-splash.html      # Initial splash screen
├── setup/                     # Setup wizard
│   ├── setup.html
│   └── setup-renderer.js
├── word/                      # Microsoft Word 97
│   ├── word.html
│   └── word-renderer.js       # 5,400+ lines
├── excel/                     # Microsoft Excel 97
│   ├── excel.html
│   └── excel-renderer.js      # 6,200+ lines
├── powerpoint/                # Microsoft PowerPoint 97
│   ├── powerpoint.html
│   ├── powerpoint-renderer.js # 10,200+ lines
│   ├── template_data.js       # Placeholder positions from .pot files
│   └── templates/             # 241 PNG backgrounds + thumbnails
├── access/                    # Microsoft Access 97
│   ├── access.html
│   └── access-renderer.js
├── outlook/                   # Microsoft Outlook 97
│   ├── outlook.html
│   └── outlook-renderer.js
├── shared/                    # Shared UI components
│   ├── win95-ui.css           # Windows 95 UI styles
│   ├── win95-ui.js            # Windows 95 UI components
│   ├── clipart/               # 200 Office 97 SVG clipart
│   ├── clipart.js             # Clip Gallery 3.0 dialog
│   └── toolbar-icons/         # Original toolbar icons
├── icons/                     # App icons (original Office 97)
└── dist/                      # Build output
```

## Lines of Code

| Component | Lines |
|-----------|-------|
| PowerPoint | ~10,200 |
| Excel | ~6,200 |
| Word | ~5,400 |
| Outlook | ~1,900 |
| Access | ~1,800 |
| **Total** | **~25,500** |

## Credits

- Original software by Microsoft Corporation (1996-1997)
- Template files from Archive.org
- Clipart from the Microsoft Office 97 Clipart Collection (Archive.org)
- Toolbar icons extracted from original Office 97 binaries

## Disclaimer

This is an educational recreation project. Microsoft Office is a trademark of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft. No proprietary code from the original Office 97 was used — all functionality was recreated from scratch.

---

*Built with mass amounts of mass in 2026*
