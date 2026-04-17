param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookJsonBase64
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'office-common.ps1')

$payload = Decode-JsonBase64 -Value $WorkbookJsonBase64
$outputPath = Ensure-ParentDirectory -Path $payload.outputPath
$sheetName = if ($payload.sheetName) { [string]$payload.sheetName } else { 'Sheet1' }
$title = if ($payload.title) { [string]$payload.title } else { 'Data Table' }
$fontName = if ($payload.fontName) { [string]$payload.fontName } else { 'Microsoft YaHei' }
$freezeHeader = [bool]$payload.freezeHeader
$leaveOpen = [bool]$payload.leaveOpen
$headerCount = @($payload.headers).Count
$rowCount = @($payload.rows).Count
$lastRow = $rowCount + 2
$lastColumnNumber = $headerCount

if ($headerCount -lt 1) {
    throw 'The workbook payload must include at least one header.'
}

$excelInfo = Get-OrNewExcelApplication
$excel = $excelInfo.App
$excel.Visible = $leaveOpen -or -not $excelInfo.Created
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

try {
    if ($sheetName.Length -gt 31) {
        $sheetName = $sheetName.Substring(0, 31)
    }
    $sheet.Name = $sheetName

    $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $lastColumnNumber)).Merge() | Out-Null
    $sheet.Cells.Item(1, 1).Value2 = $title

    for ($column = 0; $column -lt $headerCount; $column++) {
        $sheet.Cells.Item(2, $column + 1).Value2 = [string]$payload.headers[$column]
    }

    for ($row = 0; $row -lt $rowCount; $row++) {
        $values = @($payload.rows[$row])
        for ($column = 0; $column -lt $headerCount; $column++) {
            $value = if ($column -lt $values.Count) { [string]$values[$column] } else { '' }
            $sheet.Cells.Item($row + 3, $column + 1).Value2 = $value
        }
    }

    $titleRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $lastColumnNumber))
    $headerRange = $sheet.Range($sheet.Cells.Item(2, 1), $sheet.Cells.Item(2, $lastColumnNumber))
    $bodyRange = $sheet.Range($sheet.Cells.Item(3, 1), $sheet.Cells.Item($lastRow, $lastColumnNumber))
    $fullRange = $sheet.Range($sheet.Cells.Item(2, 1), $sheet.Cells.Item($lastRow, $lastColumnNumber))

    $sheet.Cells.Font.Name = $fontName
    $titleRange.Font.Size = 16
    $titleRange.Font.Bold = $true
    $titleRange.HorizontalAlignment = -4108
    $titleRange.VerticalAlignment = -4108
    $titleRange.Interior.Color = 0xF7EAD9
    $titleRange.RowHeight = 28

    $headerRange.Font.Size = 11
    $headerRange.Font.Bold = $true
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.Interior.Color = 0xD9EAF7
    $headerRange.RowHeight = 24

    $bodyRange.Font.Size = 10.5
    $bodyRange.HorizontalAlignment = -4108
    $bodyRange.VerticalAlignment = -4108
    $bodyRange.RowHeight = 22

    $fullRange.Borders.LineStyle = 1
    $fullRange.Borders.Weight = 2
    if ($excel.ActiveWindow) {
        $excel.ActiveWindow.DisplayGridlines = $false
    }

    for ($column = 1; $column -le $lastColumnNumber; $column++) {
        $sheet.Columns.Item($column).AutoFit() | Out-Null
        $width = [double]$sheet.Columns.Item($column).ColumnWidth
        if ($column -eq 1 -and $width -lt 8) {
            $sheet.Columns.Item($column).ColumnWidth = 8
        } elseif ($width -lt 10) {
            $sheet.Columns.Item($column).ColumnWidth = 10
        } elseif ($width -gt 24) {
            $sheet.Columns.Item($column).ColumnWidth = 24
        }
    }

    foreach ($cell in $bodyRange) {
        $value = [string]$cell.Value2
        if ($value -eq ([string][char]8730)) {
            $cell.Font.Color = 0x008000
            $cell.Font.Bold = $true
        } elseif ($value -eq '-') {
            $cell.Font.Color = 0x666666
        }
    }

    if ($freezeHeader) {
        $sheet.Activate() | Out-Null
        $sheet.Range('A3').Select() | Out-Null
        if ($excel.ActiveWindow) {
            $excel.ActiveWindow.FreezePanes = $true
        }
    }

    $workbook.SaveAs($outputPath, 51)

    if (-not $leaveOpen) {
        $workbook.Close($true)
        if ($excelInfo.Created) {
            $excel.Quit()
        }
    }

    Write-Output $outputPath
}
catch {
    try {
        $workbook.Close($false)
    } catch {
    }
    if ($excelInfo.Created) {
        try {
            $excel.Quit()
        } catch {
        }
    }
    throw
}
