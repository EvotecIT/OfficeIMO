#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies that customer-facing Office capability claims match generated truth.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'
$siteRootPath = (Resolve-Path -LiteralPath $SiteRoot).Path
$repositoryRootPath = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$catalogPath = Join-Path $siteRootPath 'data\office_capabilities.json'

if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
    throw "Generated capability catalog is missing: $catalogPath"
}

$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
$requirements = [ordered]@{
    word = @('.doc', '.docx', '.docm', '.dot')
    excel = @('.xls', '.xlsx', '.xlsb', '.xlsm')
    powerpoint = @('.ppt', '.pptx', '.pps', '.pot')
}

foreach ($familyId in $requirements.Keys) {
    $family = @($catalog.families | Where-Object id -eq $familyId)
    if ($family.Count -ne 1) {
        throw "Capability catalog must contain exactly one '$familyId' family."
    }

    $extensions = @($family[0].formats.extension)
    foreach ($extension in $requirements[$familyId]) {
        if ($extension -notin $extensions) {
            throw "Capability catalog family '$familyId' is missing required format '$extension'."
        }
    }

    if ([int] $family[0].contract.capabilityCount -lt 1) {
        throw "Capability catalog family '$familyId' has no tracked behaviors."
    }
}

$contractRequirements = [ordered]@{
    'OfficeIMO.Word.LegacyDoc' = 33
    'OfficeIMO.Excel.LegacyXls' = 28
    'OfficeIMO.Excel.Xlsb' = 20
    'OfficeIMO.PowerPoint.LegacyPpt' = 56
}
$contracts = @($catalog.families.contracts)
foreach ($contractId in $contractRequirements.Keys) {
    $contract = @($contracts | Where-Object id -eq $contractId)
    if ($contract.Count -ne 1) {
        throw "Capability catalog must contain exactly one '$contractId' contract."
    }
    if ([int] $contract[0].capabilityCount -ne $contractRequirements[$contractId]) {
        throw "Capability contract '$contractId' has $($contract[0].capabilityCount) behaviors; expected $($contractRequirements[$contractId])."
    }
}
if ([int] $catalog.summary.capabilityCount -ne 137) {
    throw "Capability catalog has $($catalog.summary.capabilityCount) behaviors; expected the current 137-behavior source total."
}

$stateDefinitions = @{}
foreach ($state in $catalog.fidelityStates) {
    $stateDefinitions[[string] $state.id] = $state
}
foreach ($contract in $contracts) {
    foreach ($operation in $contract.operations) {
        if ($operation.states -isnot [Array]) {
            throw "Contract '$($contract.id)' operation '$($operation.id)' must expose states as an array, including single-state operations."
        }
        foreach ($state in $operation.states) {
            if (-not $stateDefinitions.ContainsKey([string] $state.state) -or
                [string]::IsNullOrWhiteSpace([string] $state.label) -or
                [string]::IsNullOrWhiteSpace([string] $state.description)) {
                throw "Contract '$($contract.id)' operation '$($operation.id)' exposes unexplained state '$($state.state)'."
            }
        }
    }
}

$claimFiles = @(
    'data\hero.json',
    'data\format_ribbon.json',
    'data\products.json',
    'data\faq.json',
    'data\comparison.json',
    'content\products\word.md',
    'content\products\excel.md',
    'content\products\powerpoint.md',
    'content\blog\officeimo-vs-competitors.md',
    'content\blog\working-with-doc-xls-ppt-dotnet.md',
    'content\blog\choosing-document-conversion-fidelity.md',
    'content\pages\convert.md',
    'content\pages\compatibility.md',
    'content\pages\solutions.md',
    'content\conversions\doc-docx.md',
    'content\conversions\xls-xlsx.md',
    'content\conversions\xlsb-xlsx.md',
    'content\conversions\ppt-pptx.md',
    'content\solutions\legacy-office-modernization.md'
)

$claimText = ($claimFiles | ForEach-Object {
    $path = Join-Path $siteRootPath $_
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required customer-facing claim file is missing: $path"
    }
    Get-Content -LiteralPath $path -Raw
}) -join "`n"

foreach ($requiredClaim in @('DOC', 'DOCX', 'XLS', 'XLSX', 'XLSB', 'PPT', 'PPTX')) {
    if ($claimText -notmatch "(?i)(?<![a-z0-9])$requiredClaim(?![a-z0-9])") {
        throw "Customer-facing capability surfaces do not mention '$requiredClaim'."
    }
}

$forbiddenClaims = @(
    'works entirely with Open XML standards',
    'focused on COM-free Open XML workflows',
    'Are Open XML formats enough for the workload?',
    'Legacy binary Office formats.',
    '88 of the 89 production projects',
    '88 of 89 production projects',
    '85 library assemblies',
    '117|Tracked behaviors'
)
foreach ($forbiddenClaim in $forbiddenClaims) {
    if ($claimText.Contains($forbiddenClaim, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Stale or misleading customer-facing claim remains: '$forbiddenClaim'"
    }
}

$projectRequirements = [ordered]@{
    'OfficeIMO.Word\OfficeIMO.Word.csproj' = @('doc', 'docx', 'word')
    'OfficeIMO.Excel\OfficeIMO.Excel.csproj' = @('xls', 'xlsx', 'xlsb', 'excel')
    'OfficeIMO.PowerPoint\OfficeIMO.PowerPoint.csproj' = @('ppt', 'pptx', 'powerpoint')
}
foreach ($relativeProjectPath in $projectRequirements.Keys) {
    [xml] $project = Get-Content -LiteralPath (Join-Path $repositoryRootPath $relativeProjectPath) -Raw
    $tags = [string] @($project.Project.PropertyGroup.PackageTags |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -First 1)
    foreach ($tag in $projectRequirements[$relativeProjectPath]) {
        if ($tags -notmatch "(?i)(^|[;,\s])$([regex]::Escape($tag))([;,\s]|$)") {
            throw "Project '$relativeProjectPath' package tags are missing '$tag'."
        }
    }
}

$productFiles = @(Get-ChildItem -LiteralPath (Join-Path $siteRootPath 'content\products') -Filter '*.md' -File)
foreach ($productFile in $productFiles) {
    $productContent = Get-Content -LiteralPath $productFile.FullName -Raw
    foreach ($requiredMetadata in @(
        'meta.software.name:',
        'meta.software.application_category:',
        'meta.software.operating_system:',
        'meta.software.download_url:'
    )) {
        if (-not $productContent.Contains($requiredMetadata, [StringComparison]::Ordinal)) {
            throw "Product page '$($productFile.Name)' is missing '$requiredMetadata' structured metadata."
        }
    }
}

$compatibilityLayout = Get-Content -LiteralPath (Join-Path $siteRootPath 'themes\officeimo\layouts\compatibility.html') -Raw
foreach ($requiredLayoutEvidence in @(
    'family.contracts',
    'contract.hasUnimplementedCoverage',
    'state.label',
    'state.description'
)) {
    if (-not $compatibilityLayout.Contains($requiredLayoutEvidence, [StringComparison]::Ordinal)) {
        throw "Compatibility layout does not surface required contract evidence '$requiredLayoutEvidence'."
    }
}

$aotPath = Join-Path $siteRootPath 'static\data\aot-compatibility.json'
if (-not (Test-Path -LiteralPath $aotPath -PathType Leaf)) {
    throw "NativeAOT capability evidence is missing: $aotPath"
}
$aot = Get-Content -LiteralPath $aotPath -Raw | ConvertFrom-Json
$aotClaimFiles = @(
    'data\faq.json',
    'data\comparison.json',
    'content\docs\advanced\aot-trimming\index.md',
    'content\blog\aot-trimming-office.md'
)
$aotClaimText = ($aotClaimFiles | ForEach-Object {
    Get-Content -LiteralPath (Join-Path $siteRootPath $_) -Raw
}) -join "`n"
$validatedClaim = "$($aot.summary.nativeAotValidatedProjectCount) of $($aot.summary.productionProjectCount)"
if (-not $aotClaimText.Contains($validatedClaim, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Customer-facing NativeAOT claims do not match the current '$validatedClaim' project matrix."
}
$fullyRootedClaim = [string] $aot.summary.fullyRootedLibraryCount
if ($aotClaimText -notmatch "(?i)(fully roots|fully rooted|libraries are fully rooted)\D{0,20}$fullyRootedClaim|$fullyRootedClaim\D{0,20}(fully rooted|library assemblies)") {
    throw "Customer-facing NativeAOT claims do not mention the current $fullyRootedClaim fully rooted libraries."
}

$staticCatalogPath = Join-Path $siteRootPath 'static\data\office-capabilities.json'
if (-not (Test-Path -LiteralPath $staticCatalogPath -PathType Leaf)) {
    throw "Generated public capability catalog is missing: $staticCatalogPath"
}

$dataCatalogContent = Get-Content -LiteralPath $catalogPath -Raw
$staticCatalogContent = Get-Content -LiteralPath $staticCatalogPath -Raw
if (-not $dataCatalogContent.Equals($staticCatalogContent, [StringComparison]::Ordinal)) {
    throw 'The build-time and public capability catalogs do not match.'
}

$powerShellCatalogPath = Join-Path $siteRootPath 'data\pswriteoffice_command_catalog.json'
if (-not (Test-Path -LiteralPath $powerShellCatalogPath -PathType Leaf)) {
    throw "PSWriteOffice command catalog is missing: $powerShellCatalogPath"
}

$powerShellCatalog = Get-Content -LiteralPath $powerShellCatalogPath -Raw | ConvertFrom-Json
$powerShellWord = @($powerShellCatalog.families | Where-Object id -eq 'word')
$powerShellExcel = @($powerShellCatalog.families | Where-Object id -eq 'excel')
if ($powerShellWord.Count -ne 1 -or
    [string] $powerShellWord[0].description -notmatch '(?i)\bDOC\b.*\bDOCX\b') {
    throw 'PSWriteOffice Word website claims must mention both DOC and DOCX support.'
}
if ($powerShellExcel.Count -ne 1 -or
    [string] $powerShellExcel[0].description -notmatch '(?i)\bXLS\b.*\bXLSX\b') {
    throw 'PSWriteOffice Excel website claims must mention both XLS and XLSX support.'
}
if ([string] $powerShellExcel[0].description -match '(?i)\bXLSB\b') {
    throw 'PSWriteOffice Excel website claims must not advertise unsupported XLSB cmdlet workflows.'
}

$statsPath = Join-Path $siteRootPath 'data\stats.json'
if (-not (Test-Path -LiteralPath $statsPath -PathType Leaf)) {
    throw "Generated website statistics are missing: $statsPath"
}

$stats = Get-Content -LiteralPath $statsPath -Raw | ConvertFrom-Json
$statValues = @{}
foreach ($item in $stats.items) {
    $statValues[[string] $item.label] = [string] $item.value
}

if ($statValues['Tracked format behaviors'] -ne [string] $catalog.summary.capabilityCount) {
    throw 'Website statistics do not match the generated compatibility behavior count.'
}
if ($statValues['Word, Excel, and PowerPoint variants'] -ne [string] $catalog.summary.formatCount) {
    throw 'Website statistics do not match the generated Office format count.'
}

Write-Host "Capability claims verified against $($catalog.summary.formatCount) format variants and $($catalog.summary.capabilityCount) tracked behaviors."
