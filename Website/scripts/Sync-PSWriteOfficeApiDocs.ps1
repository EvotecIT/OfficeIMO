#!/usr/bin/env pwsh

[CmdletBinding()]
param(
    [string] $PSWriteOfficeRoot = $env:PSWRITEOFFICE_ROOT,
    [string] $OutputRoot = (Join-Path $PSScriptRoot '..\data\apidocs\powershell')
)

$ErrorActionPreference = 'Stop'

function Ensure-ExamplesPlaceholder {
    param(
        [Parameter(Mandatory)]
        [string] $ExamplesPath
    )

    if (-not (Test-Path -LiteralPath $ExamplesPath -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $ExamplesPath | Out-Null
    }

    $placeholderPath = Join-Path $ExamplesPath '.gitkeep'
    if (-not (Test-Path -LiteralPath $placeholderPath -PathType Leaf)) {
        Set-Content -LiteralPath $placeholderPath -Value '' -NoNewline
    }
}

function Resolve-PSWriteOfficeRoot {
    param(
        [string] $RequestedRoot
    )

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        $candidates += $RequestedRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($env:PSWRITEOFFICE_ROOT)) {
        $candidates += $env:PSWRITEOFFICE_ROOT
    }

    $websiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
    $cursor = $websiteRoot
    for ($i = 0; $i -lt 8; $i++) {
        $parent = Split-Path -Parent $cursor
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $cursor) {
            break
        }
        $candidates += (Join-Path $parent 'PSWriteOffice')
        $cursor = $parent
    }

    if ($IsWindows) {
        $candidates += 'C:\Support\GitHub\PSWriteOffice'
    }
    $candidates += '/mnt/c/Support/GitHub/PSWriteOffice'

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        try {
            $root = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        } catch {
            continue
        }

        if (Test-Path -LiteralPath (Join-Path $root 'Build\Manage-PSWriteOffice.ps1') -PathType Leaf) {
            return $root
        }
    }

    return $null
}

$resolvedRoot = Resolve-PSWriteOfficeRoot -RequestedRoot $PSWriteOfficeRoot
if (-not $resolvedRoot) {
    Ensure-ExamplesPlaceholder -ExamplesPath (Join-Path $OutputRoot 'examples')
    Write-Warning 'PSWriteOffice repo not found. Keeping the checked-in PowerShell API inputs.'
    exit 0
}

$helpCandidates = @(
    (Join-Path $resolvedRoot 'Artefacts\Unpacked\Modules\PSWriteOffice\en-US\PSWriteOffice-help.xml'),
    (Join-Path $resolvedRoot 'Sources\PSWriteOffice\bin\Release\net8.0\PSWriteOffice.dll-Help.xml'),
    (Join-Path $resolvedRoot 'Sources\PSWriteOffice\bin\Debug\net8.0\PSWriteOffice.dll-Help.xml')
)

$helpSource = $helpCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
$examplesSource = Join-Path $resolvedRoot 'Examples'

if (-not $helpSource) {
    Ensure-ExamplesPlaceholder -ExamplesPath (Join-Path $OutputRoot 'examples')
    Write-Warning "PSWriteOffice help XML not found under '$resolvedRoot'. Keeping the checked-in PowerShell API inputs."
    exit 0
}

if (-not (Test-Path -LiteralPath $examplesSource -PathType Container)) {
    Ensure-ExamplesPlaceholder -ExamplesPath (Join-Path $OutputRoot 'examples')
    Write-Warning "PSWriteOffice Examples folder not found under '$resolvedRoot'. Keeping the checked-in PowerShell API inputs."
    exit 0
}

if (-not (Test-Path -LiteralPath $OutputRoot -PathType Container)) {
    $null = New-Item -ItemType Directory -Force -Path $OutputRoot
}
$outputRootPath = (Resolve-Path -LiteralPath $OutputRoot).Path

$helpDestination = Join-Path $outputRootPath 'PSWriteOffice-Help.xml'
$examplesDestination = Join-Path $outputRootPath 'examples'

Copy-Item -LiteralPath $helpSource -Destination $helpDestination -Force

if (Test-Path -LiteralPath $examplesDestination) {
    Remove-Item -LiteralPath $examplesDestination -Recurse -Force
}
New-Item -ItemType Directory -Path $examplesDestination -Force | Out-Null
Copy-Item -Path (Join-Path $examplesSource '*') -Destination $examplesDestination -Recurse -Force
Ensure-ExamplesPlaceholder -ExamplesPath $examplesDestination

Write-Host "Synced PSWriteOffice API docs from: $resolvedRoot" -ForegroundColor Cyan
Write-Host "Help XML: $helpSource" -ForegroundColor DarkGray
Write-Host "Examples: $examplesSource" -ForegroundColor DarkGray
