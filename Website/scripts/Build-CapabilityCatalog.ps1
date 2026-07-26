#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Projects OfficeIMO compatibility contracts into website-ready capability data.

.DESCRIPTION
    Reads the generated Word, Excel, and PowerPoint compatibility contracts and
    produces a compact public catalog. The compatibility libraries remain the
    source of truth; the website receives a deterministic projection rather than
    maintaining another format-support matrix.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'

function Read-JsonFile {
    param([Parameter(Mandatory)][string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required JSON file was not found: $Path"
    }

    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)] $Value
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }

    $json = $Value | ConvertTo-Json -Depth 30
    [IO.File]::WriteAllText(
        $Path,
        $json + [Environment]::NewLine,
        [Text.UTF8Encoding]::new($false))
}

function Get-VersionPrefix {
    param([Parameter(Mandatory)][string] $ProjectPath)

    [xml] $project = Get-Content -LiteralPath $ProjectPath -Raw
    $version = @($project.Project.PropertyGroup.VersionPrefix |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -First 1)
    if ($version.Count -ne 1) {
        throw "Project '$ProjectPath' does not declare one VersionPrefix."
    }

    [string] $version[0]
}

function Get-StateCounts {
    param(
        [Parameter(Mandatory)][object[]] $Capabilities,
        [Parameter(Mandatory)][string] $Property
    )

    @($Capabilities |
        Group-Object -Property $Property |
        Sort-Object Name |
        ForEach-Object {
            $stateDefinition = $script:stateDefinitionsById[$_.Name]
            if ($null -eq $stateDefinition) {
                throw "Compatibility state '$($_.Name)' does not have a public explanation."
            }
            [ordered]@{
                state = $_.Name
                label = $stateDefinition.label
                description = $stateDefinition.description
                count = $_.Count
            }
        })
}

$siteRootPath = (Resolve-Path -LiteralPath $SiteRoot).Path
$repositoryRootPath = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$compatibilityRoot = Join-Path $repositoryRootPath 'Docs\Compatibility\generated'
$formatInventory = Read-JsonFile (Join-Path $compatibilityRoot 'office-formats.json')

$familyDefinitions = @(
    [ordered]@{
        Id = 'Word'
        Slug = 'word'
        Title = 'Microsoft Word'
        Package = 'OfficeIMO.Word'
        Project = 'OfficeIMO.Word\OfficeIMO.Word.csproj'
        Contracts = @(
            [ordered]@{
                File = 'word-legacy-doc.json'
                Label = 'DOC (Word 97–2003)'
                Scope = 'Legacy DOC import, writing, round-trip behavior, and DOC/DOCX conversion.'
            }
        )
        ProductUrl = '/products/word/'
        DocsUrl = '/docs/word/'
        ApiUrl = '/api/word/'
        Color = '#2563eb'
    },
    [ordered]@{
        Id = 'Excel'
        Slug = 'excel'
        Title = 'Microsoft Excel'
        Package = 'OfficeIMO.Excel'
        Project = 'OfficeIMO.Excel\OfficeIMO.Excel.csproj'
        Contracts = @(
            [ordered]@{
                File = 'excel-legacy-xls.json'
                Label = 'XLS (BIFF8)'
                Scope = 'Legacy XLS workbook behavior and XLS/XLSX conversion.'
            },
            [ordered]@{
                File = 'excel-xlsb.json'
                Label = 'XLSB (binary Open XML)'
                Scope = 'Binary workbook lifecycle, preservation, and XLSB/XLSX conversion.'
                OperationLabels = @{
                    create = 'Create XLSB'
                    modernToLegacy = 'XLSX to XLSB'
                    legacyToModern = 'XLSB to XLSX'
                }
            }
        )
        ProductUrl = '/products/excel/'
        DocsUrl = '/docs/excel/'
        ApiUrl = '/api/excel/'
        Color = '#059669'
    },
    [ordered]@{
        Id = 'PowerPoint'
        Slug = 'powerpoint'
        Title = 'Microsoft PowerPoint'
        Package = 'OfficeIMO.PowerPoint'
        Project = 'OfficeIMO.PowerPoint\OfficeIMO.PowerPoint.csproj'
        Contracts = @(
            [ordered]@{
                File = 'powerpoint-legacy-ppt.json'
                Label = 'PPT (PowerPoint 97–2003)'
                Scope = 'Legacy PPT import, writing, round-trip behavior, and PPT/PPTX conversion.'
            }
        )
        ProductUrl = '/products/powerpoint/'
        DocsUrl = '/docs/powerpoint/'
        ApiUrl = '/api/powerpoint/'
        Color = '#dc2626'
    }
)

$operationDefinitions = @(
    [ordered]@{ id = 'representability'; label = 'Represent'; property = 'representability' },
    [ordered]@{ id = 'import'; label = 'Read and import'; property = 'legacyImport' },
    [ordered]@{ id = 'create'; label = 'Create legacy'; property = 'newLegacyWrite' },
    [ordered]@{ id = 'roundTrip'; label = 'Edit and round-trip'; property = 'legacyRoundTrip' },
    [ordered]@{ id = 'modernToLegacy'; label = 'Modern to legacy'; property = 'modernToLegacy' },
    [ordered]@{ id = 'legacyToModern'; label = 'Legacy to modern'; property = 'legacyToModern' }
)

$stateDefinitions = @(
    [ordered]@{ id = 'Native'; label = 'Native'; description = 'Represented directly in the destination format.' },
    [ordered]@{ id = 'Approximated'; label = 'Editable approximation'; description = 'Retained as editable content with a documented approximation.' },
    [ordered]@{ id = 'Approximation'; label = 'Closest representation'; description = 'Mapped to the closest supported representation, with a reported fidelity difference.' },
    [ordered]@{ id = 'Rasterized'; label = 'Visual fallback'; description = 'Appearance retained as a static image when editability is unavailable.' },
    [ordered]@{ id = 'PreservedOpaque'; label = 'Preserved records'; description = 'Original records retained without claiming full editing semantics.' },
    [ordered]@{ id = 'Opaque'; label = 'Opaque content'; description = 'Content carried without interpreting its internal feature semantics.' },
    [ordered]@{ id = 'EmbeddedSource'; label = 'Source retained'; description = 'Original source embedded with hash verification for recovery.' },
    [ordered]@{ id = 'Blocked'; label = 'Blocked'; description = 'Operation refused when proceeding would silently misrepresent the result.' },
    [ordered]@{ id = 'NotApplicable'; label = 'Not applicable'; description = 'The operation does not apply to this feature or format direction.' },
    [ordered]@{ id = 'NotImplemented'; label = 'Not implemented'; description = 'The contract explicitly records a path that is not implemented yet.' },
    [ordered]@{ id = 'NotRepresentable'; label = 'No destination equivalent'; description = 'The destination format has no representation for this feature.' },
    [ordered]@{ id = 'Dropped'; label = 'Dropped with diagnostics'; description = 'The feature is omitted and reported as known loss.' }
)
$script:stateDefinitionsById = @{}
foreach ($stateDefinition in $stateDefinitions) {
    $script:stateDefinitionsById[$stateDefinition.id] = $stateDefinition
}

$families = foreach ($definition in $familyDefinitions) {
    $formatFamily = @($formatInventory.families |
        Where-Object { $_.id -eq $definition.Id })
    if ($formatFamily.Count -ne 1) {
        throw "Format inventory must contain exactly one '$($definition.Id)' family."
    }

    $formats = @($formatFamily[0].formats | ForEach-Object {
        [ordered]@{
            id = $_.id
            extension = $_.extension
            generation = $_.generation
            documentKind = $_.documentKind
            encoding = $_.encoding
            macroEnabled = [bool] $_.isMacroEnabled
        }
    })

    $allCapabilities = @()
    $contractProjections = @()
    foreach ($contractDefinition in $definition.Contracts) {
        $contract = Read-JsonFile (Join-Path $compatibilityRoot $contractDefinition.File)
        $capabilities = @($contract.capabilities)
        if ($capabilities.Count -eq 0) {
            throw "Compatibility contract '$($contractDefinition.File)' contains no capabilities."
        }

        $allCapabilities += $capabilities
        $formatIds = @($capabilities.formatId | Sort-Object -Unique)
        $contractFormats = @($formats |
            Where-Object { $_.id -in $formatIds } |
            ForEach-Object extension)
        $contractOperations = @($operationDefinitions | ForEach-Object {
            $operationLabel = $_.label
            if ($contractDefinition.Contains('OperationLabels') -and
                $contractDefinition.OperationLabels.ContainsKey($_.id)) {
                $operationLabel = $contractDefinition.OperationLabels[$_.id]
            }
            [ordered]@{
                id = $_.id
                label = $operationLabel
                states = @(Get-StateCounts -Capabilities $capabilities -Property $_.property)
            }
        })

        $contractProjections += [ordered]@{
            id = $contract.id
            label = $contractDefinition.Label
            scope = $contractDefinition.Scope
            source = $contractDefinition.File
            schemaVersion = $contract.schemaVersion
            formats = $contractFormats
            capabilityCount = $capabilities.Count
            categoryCount = @($capabilities.category | Sort-Object -Unique).Count
            hasUnimplementedCoverage = [bool] $contract.hasUnimplementedCoverage
            operations = $contractOperations
        }
    }

    $operations = @($operationDefinitions | ForEach-Object {
        [ordered]@{
            id = $_.id
            label = $_.label
            states = @(Get-StateCounts -Capabilities $allCapabilities -Property $_.property)
        }
    })

    $legacyFormats = @($formats |
        Where-Object generation -eq 'Legacy' |
        ForEach-Object extension)
    $modernFormats = @($formats |
        Where-Object generation -eq 'Modern' |
        ForEach-Object extension)

    [ordered]@{
        id = $definition.Slug
        title = $definition.Title
        package = $definition.Package
        packageVersion = Get-VersionPrefix (Join-Path $repositoryRootPath $definition.Project)
        productUrl = $definition.ProductUrl
        docsUrl = $definition.DocsUrl
        apiUrl = $definition.ApiUrl
        color = $definition.Color
        legacyFormats = $legacyFormats
        modernFormats = $modernFormats
        formats = $formats
        contract = [ordered]@{
            capabilityCount = $allCapabilities.Count
            categoryCount = @($allCapabilities.category | Sort-Object -Unique).Count
            hasUnimplementedCoverage = @($contractProjections | Where-Object hasUnimplementedCoverage).Count -gt 0
            operations = $operations
        }
        contracts = $contractProjections
    }
}

$capabilityCount = [int] (($families |
    ForEach-Object { $_.contract.capabilityCount } |
    Measure-Object -Sum).Sum)
$formatCount = [int] (($families |
    ForEach-Object { $_.formats.Count } |
    Measure-Object -Sum).Sum)

$catalog = [ordered]@{
    schemaVersion = 1
    source = [ordered]@{
        formatInventory = 'Docs/Compatibility/generated/office-formats.json'
        compatibilityContracts = @(
            'Docs/Compatibility/generated/word-legacy-doc.json',
            'Docs/Compatibility/generated/excel-legacy-xls.json',
            'Docs/Compatibility/generated/excel-xlsb.json',
            'Docs/Compatibility/generated/powerpoint-legacy-ppt.json'
        )
        claim = 'Feature-level compatibility with explicit native, approximation, visual, preservation, and blocked states.'
    }
    summary = [ordered]@{
        familyCount = $families.Count
        formatCount = $formatCount
        capabilityCount = $capabilityCount
    }
    fidelityStates = $stateDefinitions
    families = @($families)
}

$dataPath = Join-Path $siteRootPath 'data\office_capabilities.json'
$staticPath = Join-Path $siteRootPath 'static\data\office-capabilities.json'
Write-JsonFile -Path $dataPath -Value $catalog
Write-JsonFile -Path $staticPath -Value $catalog

$documentationCatalog = Read-JsonFile (Join-Path $siteRootPath 'data\documentation_catalog.json')
$powerShellCatalog = Read-JsonFile (Join-Path $siteRootPath 'data\pswriteoffice_command_catalog.json')
$stats = [ordered]@{
    items = @(
        [ordered]@{ value = [string] @($documentationCatalog.components).Count; label = 'Production components' },
        [ordered]@{ value = [string] $powerShellCatalog.module.commandCount; label = 'PowerShell commands' },
        [ordered]@{ value = [string] $capabilityCount; label = 'Tracked format behaviors' },
        [ordered]@{ value = [string] $formatCount; label = 'Word, Excel, and PowerPoint variants' }
    )
}
Write-JsonFile -Path (Join-Path $siteRootPath 'data\stats.json') -Value $stats

Write-Host "Generated capability catalog: $($families.Count) families, $formatCount format variants, $capabilityCount tracked behaviors."
