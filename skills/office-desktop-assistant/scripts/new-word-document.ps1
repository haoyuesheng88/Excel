param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentJsonBase64
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'office-common.ps1')

$payload = Decode-JsonBase64 -Value $DocumentJsonBase64
$outputPath = Ensure-ParentDirectory -Path $payload.outputPath
$fontName = if ($payload.fontName) { [string]$payload.fontName } else { 'Microsoft YaHei' }
$bodyFontSize = if ($payload.bodyFontSize) { [double]$payload.bodyFontSize } else { 11 }
$titleFontSize = if ($payload.titleFontSize) { [double]$payload.titleFontSize } else { 16 }
$openAfterSave = [bool]$payload.openAfterSave

$wordInfo = Get-OrNewWordApplication
$word = $wordInfo.App
$word.Visible = $openAfterSave -or -not $wordInfo.Created
$word.DisplayAlerts = 0
$document = $word.Documents.Add()

try {
    if ($payload.title) {
        $selection = $word.Selection
        Set-WordSelectionFont -Selection $selection -FontName $fontName -FontSize $titleFontSize -Bold
        $selection.TypeText([string]$payload.title)
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    }

    foreach ($paragraph in @($payload.paragraphs)) {
        $selection = $word.Selection
        Set-WordSelectionFont -Selection $selection -FontName $fontName -FontSize $bodyFontSize
        $selection.TypeText([string]$paragraph)
        $selection.TypeParagraph()
    }

    $document.SaveAs2($outputPath, 16)

    if (-not $openAfterSave) {
        $document.Close($false)
        if ($wordInfo.Created) {
            $word.Quit()
        }
    }

    Write-Output $outputPath
}
catch {
    try {
        $document.Close($false)
    } catch {
    }
    if ($wordInfo.Created) {
        try {
            $word.Quit()
        } catch {
        }
    }
    throw
}
