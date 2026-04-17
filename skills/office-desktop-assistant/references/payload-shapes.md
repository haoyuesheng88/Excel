# Payload Shapes

Use these JSON shapes before UTF-8 + base64 encoding.

## Word Document Payload

```json
{
  "outputPath": "C:\\Users\\name\\Desktop\\example.docx",
  "title": "2026-04-18 Shanghai Fengxian Weather",
  "paragraphs": [
    "Current conditions are cloudy, around 18 C.",
    "Add a Japanese translation on the next line if requested."
  ],
  "fontName": "Microsoft YaHei",
  "bodyFontSize": 11,
  "titleFontSize": 16,
  "openAfterSave": true
}
```

## Excel Workbook Payload

```json
{
  "outputPath": "C:\\Users\\name\\Desktop\\permissions.xlsx",
  "sheetName": "Permissions",
  "title": "Permission Matrix",
  "headers": ["No.", "Permission", "Admin", "Engineer"],
  "rows": [
    ["1.", "User Management", "鈭?, "-"],
    ["2.", "System Exit", "鈭?, "鈭?]
  ],
  "fontName": "Microsoft YaHei",
  "freezeHeader": true,
  "leaveOpen": true
}
```
