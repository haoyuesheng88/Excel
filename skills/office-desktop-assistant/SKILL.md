---
name: office-desktop-assistant
description: Automate Microsoft Word and Excel desktop work on Windows. Use when Codex needs to type into an already open Word document, create a formatted docx on disk, create or style an Excel workbook, transcribe structured image content into Excel, or prepare Office-ready weather notes and translated follow-up lines for Word.
---

# Office Desktop Assistant

Use this skill for practical Microsoft Office desktop work on Windows.

Prefer the bundled PowerShell scripts over ad hoc COM automation when the task matches one of the packaged flows.

## Choose The Workflow

1. Insert text into an already open Word document:
Run [scripts/insert-into-active-word.ps1](./scripts/insert-into-active-word.ps1).

2. Create a new formatted `.docx` file:
Run [scripts/new-word-document.ps1](./scripts/new-word-document.ps1).

3. Create a new formatted `.xlsx` table or matrix:
Run [scripts/new-excel-table.ps1](./scripts/new-excel-table.ps1).

4. The user needs current weather in Word:
Look up live weather first, draft the text, then use either the active-Word script or the new-docx script depending on whether the user wants an open document or a saved file.

## Encoding Rule

Do not pass long non-ASCII text directly on the command line when you can avoid it.

Instead:

1. Build the text or JSON payload in PowerShell.
2. Encode it as UTF-8.
3. Convert it to base64.
4. Pass the base64 string to the script.

Use this pattern:

```powershell
$json = @{ title = 'Example'; paragraphs = @('Line 1', 'Line 2') } | ConvertTo-Json -Depth 10
$payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
& ".\skills\office-desktop-assistant\scripts\new-word-document.ps1" -DocumentJsonBase64 $payload
```

This avoids the Windows code-page problems that can turn Chinese text into `?`.

## Word Rules

- If the user says Word is already open, prefer `insert-into-active-word.ps1`.
- Keep inserted text concise unless the user asks for a fuller note.
- Use an absolute date when the request says `today`, `tomorrow`, or similar.
- For translations, translate the already drafted content unless the user explicitly asks for a fresh lookup.
- If the user asks for Microsoft YaHei, pass `Microsoft YaHei` as the font name.

For the Word payload shape, see [references/payload-shapes.md](./references/payload-shapes.md).

## Excel Rules

- Structure the data first: `title`, `sheetName`, `headers`, and `rows`.
- For screenshot-to-table tasks, transcribe the content carefully before opening Excel.
- Prefer clean presentation: title row, bold headers, centered cells, borders, and reasonable widths.
- Leave the workbook open when the user asks to open Excel or continue editing there.
- If the data is a permission matrix or status table, keep checkmarks visually distinct from placeholders such as `-`.

For the Excel payload shape, see [references/payload-shapes.md](./references/payload-shapes.md).

## Validation

After running a script, verify:

- the file path exists when the task should save a file
- Word or Excel is open when the user asked for it
- the document contains readable text and the workbook has the expected title, headers, and rows

## Deliverable

Report the saved path or confirm that the content was inserted into the active Office window.
