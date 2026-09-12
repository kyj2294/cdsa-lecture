param(
    [string]$ClaudeHome = (Join-Path $env:USERPROFILE '.claude')
)

$ErrorActionPreference = 'Stop'
$pluginRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$pluginsDirectory = Join-Path $ClaudeHome 'plugins'
$skillsDirectory = Join-Path $ClaudeHome 'skills'
$pluginLink = Join-Path $pluginsDirectory 'cdsappt'
$skillLink = Join-Path $skillsDirectory 'cdsappt'

New-Item -ItemType Directory -Path $pluginsDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $skillsDirectory -Force | Out-Null

function Ensure-Junction([string]$Path, [string]$Target) {
    $resolvedTarget = (Resolve-Path -LiteralPath $Target).Path
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if ($null -ne $item) {
        $currentTarget = @($item.Target)[0]
        if ($item.LinkType -eq 'Junction') {
            $resolvedCurrentTarget = if ($currentTarget) {
                Resolve-Path -LiteralPath $currentTarget -ErrorAction SilentlyContinue
            }
            if ($resolvedCurrentTarget -and $resolvedCurrentTarget.Path -eq $resolvedTarget) {
                return
            }
            $item.Delete()
        }
        else {
            throw "Refusing to replace existing non-junction path: $Path"
        }
    }
    New-Item -ItemType Junction -Path $Path -Target $resolvedTarget | Out-Null
}

Ensure-Junction $pluginLink $pluginRoot
Ensure-Junction $skillLink (Join-Path $pluginLink 'skills\cdsappt')

[pscustomobject]@{
    Plugin = $pluginLink
    Skill = $skillLink
    Source = $pluginRoot
}
