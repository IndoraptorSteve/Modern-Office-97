Microsoft Office 97 -- Electron Recreation

A faithful recreation of Microsoft Office 97 Professional Edition, built with Electron. Every detail -- from the installer wizard to the toolbar icons -- matches the original 1997 release.

---

SETUP WIZARD
Authentic multi-step installer (CD Key, Name/Org, Install Type, folder path, disk space)
File copy progress with rotating feature ads on teal background
Puzzle piece animation

MICROSOFT WORD 97
Full document editor with interactive ruler (draggable margins, indents, tab stops)
Font/size dropdowns, Bold/Italic/Underline, Find & Replace
Page Setup, Normal and Page Layout views, Document Map
Open/save .docx files, print support

MICROSOFT EXCEL 97
500 rows x 104 columns (A through CZ)
Formula engine (SUM, AVERAGE, IF, VLOOKUP, etc.)
Chart Wizard (4-step, bar/line/pie/scatter/area)
Sort, AutoFill, Format Cells, Goal Seek, freeze panes

MICROSOFT POWERPOINT 97
242 design templates from original .pot files (real assets)
Click-and-drag shape drawing (line, arrow, rect, oval, text box)
30+ AutoShapes (Basic, Block Arrows, Flowchart, Stars, Callouts)
Fill Effects: 24 gradient presets, 24 textures, 48 hatch patterns
3-D Effects: 20 presets with depth/direction/lighting
Shadow Effects: 10 preset styles
200 original Office 97 clipart via Clip Gallery 3.0
Chart Wizard, table editor, slide transitions
Slide Sorter, Notes pane, Outline view, slideshow mode
10,200+ lines of code

MICROSOFT ACCESS 97
Table datasheet + design view, form view, query builder with SQL
Import/Export CSV, sample database templates

MICROSOFT OUTLOOK 97
Inbox with sample emails, Calendar (day/month), Contacts, Tasks
Compose/Reply/Forward, Services setup wizard
---

UI AUTHENTICITY
Custom Win95 UI framework (buttons, dialogs, toolbars, menus, scrollbars)
Original toolbar icons extracted from real Office 97
Pixelated font rendering for period-accurate look
Office 97 standard 40-color picker throughout

TECH STACK
Electron 28 + Vanilla JavaScript (no frameworks, just like the 90s)
mammoth/docx (Word), SheetJS (Excel), pptxgenjs (PowerPoint)

LINES OF CODE
PowerPoint: ~10,200
Excel: ~6,200
Word: ~5,400
Outlook: ~1,900
Access: ~1,800
Total: ~25,500

BUILD
npm install
npm start          # run in dev
npm run build      # build Windows .exe
