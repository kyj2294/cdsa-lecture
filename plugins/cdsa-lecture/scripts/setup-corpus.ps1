param(
    [Parameter(Mandatory = $true)]
    [string]$Source,
    [string]$Database = (Join-Path $env:LOCALAPPDATA "cdsa-lecture\lecture-library.sqlite"),
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$sourcePath = (Resolve-Path -LiteralPath $Source).Path
$dbPath = [System.IO.Path]::GetFullPath($Database)
$dbDirectory = Split-Path -Parent $dbPath
New-Item -ItemType Directory -Force -Path $dbDirectory | Out-Null

$skillRoot = Join-Path $PSScriptRoot "..\skills\cdsappt"
$indexer = Join-Path $skillRoot "scripts\lecture_library.py"

& $Python $indexer --db $dbPath build --source $sourcePath
if ($LASTEXITCODE -ne 0) {
    throw "Corpus indexing failed with exit code $LASTEXITCODE"
}

$store = Join-Path $skillRoot "scripts\slide_store.py"
& $Python $store --db $dbPath ingest
if ($LASTEXITCODE -ne 0) {
    throw "Corpus embedding failed with exit code $LASTEXITCODE"
}

Write-Host "Corpus database: $dbPath"
Write-Host "Set it for this PowerShell session with:"
Write-Host ('$env:CDSA_LECTURE_DB = "' + $dbPath + '"')
