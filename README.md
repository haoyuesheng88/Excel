# Office Skills Pack

Reusable Codex skill bundle for Microsoft Office desktop automation on Windows.

This repository packages portable Office-related skills:

- `office-desktop-assistant`: automate Word and Excel desktop tasks, including:
  - insert text into the active Word document
  - create formatted `.docx` files
  - create formatted `.xlsx` tables
  - transcribe structured image content into Excel
  - prepare current-weather notes for Word when a live lookup is needed
- `excel-psychrometric-tool`: build a macro-enabled Excel moist-air calculator workbook that accepts any two independent inputs among air temperature, relative humidity, humidity ratio, and dew point temperature, then calculates the remaining two with red input cells and green output cells

## Requirements

- Windows
- PowerShell 5.1 or later
- Desktop Microsoft Word and/or Excel installed for the tasks you want to run
- Codex desktop or another Codex environment that loads skills from `$CODEX_HOME/skills` or `~/.codex/skills`

## Install On Another PC

Clone the repository and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\install.ps1
```

By default, the script installs every bundled skill into:

- `$env:CODEX_HOME\skills` when `CODEX_HOME` is set
- otherwise `~/.codex/skills`

To install only one skill:

```powershell
.\install.ps1 -SkillName excel-psychrometric-tool
```

## Skill Location

The installed skill folders are:

```text
skills/excel-psychrometric-tool
skills/office-desktop-assistant
```

## Example Prompts

- `Use $office-desktop-assistant to type today's Shanghai Fengxian weather into the open Word document.`
- `Use $office-desktop-assistant to create a formatted docx on the desktop with Microsoft YaHei font.`
- `Use $office-desktop-assistant to put the table from this screenshot into Excel and make it presentation-ready.`
- `Use $office-desktop-assistant to build a clean Excel workbook from these headers and rows and leave it open.`
- `Use $excel-psychrometric-tool to build or refresh an Excel moist-air calculator workbook.`
- `Use $excel-psychrometric-tool to create an Excel workbook that calculates humidity ratio and dew point from air temperature and relative humidity.`
