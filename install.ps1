param(
    [string]$SkillName = ''
)

$ErrorActionPreference = 'Stop'

if ($env:CODEX_HOME) {
    $skillRoot = Join-Path $env:CODEX_HOME 'skills'
} else {
    $skillRoot = Join-Path $HOME '.codex\skills'
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
New-Item -ItemType Directory -Force -Path $skillRoot | Out-Null

if ([string]::IsNullOrWhiteSpace($SkillName)) {
    $sources = Get-ChildItem -LiteralPath (Join-Path $repoRoot 'skills') -Directory | Sort-Object Name
    if (-not $sources) {
        throw "No bundled skills found under $(Join-Path $repoRoot 'skills')"
    }

    foreach ($sourceDir in $sources) {
        $target = Join-Path $skillRoot $sourceDir.Name
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $sourceDir.FullName -Destination $target -Recurse -Force
        Write-Output "Installed $($sourceDir.Name) to $target"
    }
} else {
    $source = Join-Path $repoRoot "skills\$SkillName"
    $target = Join-Path $skillRoot $SkillName

    if (-not (Test-Path -LiteralPath $source)) {
        throw "Skill source not found: $source"
    }

    Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $source -Destination $target -Recurse -Force

    Write-Output "Installed $SkillName to $target"
}
