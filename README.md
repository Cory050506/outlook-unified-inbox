# Outlook Unified Inbox

A lightweight VBA macro for **Outlook Classic on Windows** that creates a practical unified inbox by automatically copying incoming mail from multiple account inboxes into one real Outlook folder.

Instead of relying on a search view or Outlook’s built-in aggregate views, this project uses a real folder that you choose and keeps it populated with incoming messages from all watched inboxes.

## Features

- Watches the Inbox of each Outlook account/store
- Copies new incoming messages into a single chosen Unified Inbox folder
- Uses a **real Outlook folder**, not just a search or filtered view
- Runs entirely inside Outlook VBA
- Lightweight and easy to customize
- Optional metadata-based sync logic between source items and unified copies
- Supports delete syncing between the source inbox and the unified folder for newly linked items

## Why this exists

Outlook Classic does not offer a simple true unified inbox folder across multiple accounts in the way many users expect. This project works around that limitation by copying messages into a real destination folder that can be pinned, sorted, searched, and used like a normal inbox.

## How it works

The macro attaches Outlook item event handlers to:

- each source Inbox
- the selected Unified Inbox folder

When a new message arrives in a watched source inbox, the macro:

1. detects the new mail item
2. creates a copy
3. moves the copy into the chosen Unified Inbox folder
4. stores hidden metadata that links the copy to the original item
5. optionally stores reverse-link metadata on the original item

When a linked message is deleted, the macro compares folder snapshots to determine which item was removed and attempts to delete the matching counterpart.

## Requirements

- **Microsoft Outlook Classic for Windows**
- VBA/macros enabled
- A real Outlook folder to use as the Unified Inbox
- One or more configured Outlook accounts/stores

## Project structure

- `ThisOutlookSession`  
  Starts the macro when Outlook launches.

- `modUnified`  
  Contains the main logic for:
  - startup
  - folder selection
  - syncing
  - metadata storage
  - snapshot rebuilding
  - delete handling

- `clsFolderWatcher`  
  Watches inbox and unified-folder item events.

## Installation

1. Download the ZIP of this repository (Code -> Download ZIP) and extract it to a folder on your computer.
2. Open **Outlook Classic**.
3. Press `Alt + F11` to open the VBA editor.
4. In the VBA editor:
   - Choose **File - Import File**
   - Import `modUnified.bas`
   - Import `clsFolderWatcher.cls`
5. Open **ThisOutlookSession** in the Outlook VBA project and paste in the contents from `ThisOutlookSession`. (If you can't see the Project Explorer, choose View - Project Explorer and it should appear on the left side of the window)
6. Save the VBA project.
7. Close and restart Outlook.
8. When prompted, choose the real folder you want to use as your Unified Inbox.

## Notes about importing

- `modUnified.bas` is a standard module and should be imported with **Import File...**
- `clsFolderWatcher.cls` is a class module and should also be imported with **Import File...**
- `ThisOutlookSession` is not imported the same way; open the existing **ThisOutlookSession** object in Outlook and paste the code into it manually

## Resetting the selected unified folder

If you want Outlook to ask you to choose the unified folder again, run this in the Immediate Window:

Notes
This project is for Outlook Classic, not New Outlook.
It works best with new messages processed after setup.
Older messages that were copied before metadata linking was added may not support full two-way delete syncing.
Bulk deletes may not behave perfectly because Outlook VBA delete events are limited and do not directly identify the removed item.
This is a VBA macro project, not a COM add-in.

Known limitations

- Delete syncing depends on Outlook item events and stored metadata.
- Existing older copied items may not be fully linked.
- Behavior can vary by account type and store type.
- Very large or rapid mailbox changes may require rebuilding snapshots.


Customization ideas

This project is intentionally simple and can be customized to:

- exclude certain accounts
- exclude certain folders
- filter by sender or subject
- skip duplicates
- log activity more aggressively
- add better error handling
- expand sync behavior beyond delete handling


License

MIT License

Copyright (c) 2026

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
