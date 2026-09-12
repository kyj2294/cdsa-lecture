param([Parameter(Mandatory=$true)][string]$Run)
$ErrorActionPreference = 'Stop'
$runRoot = (Resolve-Path -LiteralPath $Run).Path
$sourceDeck = Join-Path $runRoot 'candidate.pptx'
$before = (Get-FileHash -LiteralPath $sourceDeck -Algorithm SHA256).Hash.ToLowerInvariant()
$finalOutput = Join-Path $runRoot 'renders'
New-Item -ItemType Directory -Path $finalOutput -Force | Out-Null

$workingDeck = $sourceDeck
$workingOutput = $finalOutput
$renderTemp = $null
if ($sourceDeck.Length -gt 180 -or $finalOutput.Length -gt 180) {
    $tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $renderTemp = Join-Path $tempBase ('cdsappt-' + [guid]::NewGuid().ToString('N').Substring(0,8))
    New-Item -ItemType Directory -LiteralPath $renderTemp | Out-Null
    $workingDeck = Join-Path $renderTemp 'candidate.pptx'
    $workingOutput = Join-Path $renderTemp 'renders'
    New-Item -ItemType Directory -LiteralPath $workingOutput | Out-Null
    Copy-Item -LiteralPath $sourceDeck -Destination $workingDeck
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$deck = $null
try {
    $deck = $powerPoint.Presentations.Open($workingDeck, -1, 0, 0)
    $pages = @()
    for ($page = 1; $page -le $deck.Slides.Count; $page++) {
        $name = 'page-{0:D3}.png' -f $page
        $workingDestination = Join-Path $workingOutput $name
        $finalDestination = Join-Path $finalOutput $name
        $deck.Slides.Item($page).Export($workingDestination, 'PNG', 1600, 900)
        if ($workingOutput -ne $finalOutput) {
            Copy-Item -LiteralPath $workingDestination -Destination $finalDestination -Force
        }
        $pages += @{ page=$page; render=('renders/' + $name); sha256=(Get-FileHash -LiteralPath $finalDestination -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
    $after = (Get-FileHash -LiteralPath $sourceDeck -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($before -ne $after) { throw 'PPTX changed while rendering; render again.' }
    @{ engine='Microsoft PowerPoint'; deck_sha256=$before; pages=$pages } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $finalOutput 'manifest.json') -Encoding utf8
} finally {
    if ($null -ne $deck) { $deck.Saved=-1; $deck.Close() }
    # Do not quit the user's existing PowerPoint application.
    if ($null -ne $renderTemp -and (Test-Path -LiteralPath $renderTemp)) {
        $resolvedTemp = (Resolve-Path -LiteralPath $renderTemp).Path
        $tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if ($resolvedTemp.StartsWith($tempBase,[System.StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $resolvedTemp).StartsWith('cdsappt-')) {
            Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
        }
    }
}
