function Decode-Utf8Base64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
}

function Decode-JsonBase64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $json = Decode-Utf8Base64 -Value $Value
    return $json | ConvertFrom-Json
}

function Ensure-ParentDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $parent = Split-Path -Parent $fullPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    return $fullPath
}

function Get-OrNewWordApplication {
    try {
        return @{
            App = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            Created = $false
        }
    } catch {
        return @{
            App = New-Object -ComObject Word.Application
            Created = $true
        }
    }
}

function Get-ActiveWordApplication {
    return [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
}

function Get-OrNewExcelApplication {
    try {
        return @{
            App = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            Created = $false
        }
    } catch {
        return @{
            App = New-Object -ComObject Excel.Application
            Created = $true
        }
    }
}

function Set-WordSelectionFont {
    param(
        [Parameter(Mandatory = $true)]
        $Selection,
        [string]$FontName = 'Microsoft YaHei',
        [double]$FontSize = 11,
        [switch]$Bold
    )

    $Selection.Font.Name = $FontName
    $Selection.Font.NameFarEast = $FontName
    $Selection.Font.Size = $FontSize
    $Selection.Font.Bold = [int][bool]$Bold
}
