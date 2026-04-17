param(
    [Parameter(Mandatory = $true)]
    [string]$TextBase64,

    [switch]$NewParagraph,
    [string]$FontName = 'Microsoft YaHei',
    [double]$FontSize = 11,
    [switch]$Bold
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'office-common.ps1')

$word = Get-ActiveWordApplication
$text = Decode-Utf8Base64 -Value $TextBase64
$selection = $word.Selection

if ($NewParagraph) {
    $selection.TypeParagraph()
}

Set-WordSelectionFont -Selection $selection -FontName $FontName -FontSize $FontSize -Bold:$Bold
$selection.TypeText($text)

Write-Output 'Inserted text into the active Word document.'
