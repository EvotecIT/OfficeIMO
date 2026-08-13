[CmdletBinding()]
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $CatalogPath = '',
    [string] $JsonOutputPath = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path $RepositoryRoot 'Website\data\documentation_catalog.json'
}

$libraryHostPath = Join-Path $RepositoryRoot 'OfficeIMO.All.AotSmoke\OfficeIMO.All.AotSmoke.csproj'
[xml] $libraryHost = Get-Content -LiteralPath $libraryHostPath -Raw

$referencedLibraries = @($libraryHost.Project.ItemGroup.ProjectReference |
    ForEach-Object { [string] $_.Include } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Replace('\', '/')) } |
    Sort-Object -Unique)
$rootedLibraries = @($libraryHost.Project.ItemGroup.TrimmerRootAssembly |
    ForEach-Object { [string] $_.Include } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object -Unique)
$securityHostPath = Join-Path $RepositoryRoot 'OfficeIMO.Security.AotSmoke\OfficeIMO.Security.AotSmoke.csproj'
[xml] $securityHost = Get-Content -LiteralPath $securityHostPath -Raw
$dedicatedRootedLibraries = @($securityHost.Project.ItemGroup.ProjectReference |
    ForEach-Object { [string] $_.Include } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Replace('\', '/')) } |
    Sort-Object -Unique)
$fullyRootedLibraries = @($rootedLibraries + $dedicatedRootedLibraries | Sort-Object -Unique)
$boundedLibraries = @($referencedLibraries | Where-Object { $_ -notin $rootedLibraries })

$nativeTools = @(
    [ordered]@{
        name = 'OfficeIMO.Tool'
        evidence = 'The unified native executable starts and exposes namespaced HTML, Reader, and Markup commands.'
    }
)
$managedOnly = @(
    [ordered]@{
        name = 'OfficeIMO.Html.Pdf.Browser'
        classification = 'managed-cross-platform'
        evidence = 'The optional Chromium bridge builds and executes as managed code on supported operating systems; its HtmlTinkerX and Playwright browser runtime is not advertised as NativeAOT-compatible.'
    }
    [ordered]@{
        name = 'OfficeIMO.MarkdownRenderer.Wpf'
        classification = 'managed-windows'
        evidence = 'WPF executable publishing rejects trimming with NETSDK1168; validate this UI package with the managed Windows test lane.'
    }
)

if ($dedicatedRootedLibraries.Count -ne 1 -or $dedicatedRootedLibraries[0] -ne 'OfficeIMO.Security') {
    throw "The optional security NativeAOT host must root exactly OfficeIMO.Security; found $($dedicatedRootedLibraries -join ', ')."
}
if ($fullyRootedLibraries.Count -ne 95) {
    throw "Expected 95 fully rooted production libraries across the ordinary and optional-security hosts, found $($fullyRootedLibraries.Count)."
}
if ($boundedLibraries.Count -ne 1 -or $boundedLibraries[0] -ne 'OfficeIMO.GoogleWorkspace.Auth.GoogleApis') {
    throw "The bounded NativeAOT library set changed: $($boundedLibraries -join ', ')."
}

$catalog = Get-Content -LiteralPath $CatalogPath -Raw | ConvertFrom-Json
$productionNames = @($catalog.components.name | Sort-Object -Unique)
$classifiedNames = @(
    $fullyRootedLibraries
    $boundedLibraries
    $nativeTools.name
    $managedOnly.name
) | Sort-Object -Unique

$missing = @(Compare-Object -ReferenceObject $productionNames -DifferenceObject $classifiedNames |
    Where-Object SideIndicator -EQ '<=' |
    ForEach-Object InputObject)
$unexpected = @(Compare-Object -ReferenceObject $productionNames -DifferenceObject $classifiedNames |
    Where-Object SideIndicator -EQ '=>' |
    ForEach-Object InputObject)
if ($missing.Count -gt 0 -or $unexpected.Count -gt 0) {
    throw "NativeAOT coverage does not match the production catalog. Missing: $($missing -join ', '); unexpected: $($unexpected -join ', ')."
}

$components = foreach ($component in @($catalog.components | Sort-Object name)) {
    $name = [string] $component.name
    if ($name -in $rootedLibraries) {
        $classification = 'native-full-surface'
        $nativeValidated = $true
        $evidence = 'The complete assembly is rooted in the cross-platform NativeAOT host, compiled into native code, and the host starts successfully.'
    } elseif ($name -in $dedicatedRootedLibraries) {
        $classification = 'native-full-surface'
        $nativeValidated = $true
        $evidence = 'The optional security provider is rooted in its dedicated NativeAOT host, then executes CMS and XML DSig signing and verification without changing the ordinary format-consumer graph.'
    } elseif ($name -eq 'OfficeIMO.GoogleWorkspace.Auth.GoogleApis') {
        $classification = 'native-bounded-workflow'
        $nativeValidated = $true
        $evidence = 'The token-store adapter round-trips from the native host. Google authorization APIs remain subject to Google.Apis and Newtonsoft.Json trimming warnings when the entire dependency is rooted.'
    } elseif ($name -in $nativeTools.name) {
        $classification = 'native-executable'
        $nativeValidated = $true
        $evidence = [string] ($nativeTools | Where-Object name -EQ $name).evidence
    } elseif ($name -in $managedOnly.name) {
        $managedEntry = @($managedOnly | Where-Object name -EQ $name)[0]
        $classification = [string] $managedEntry.classification
        $nativeValidated = $false
        $evidence = [string] $managedEntry.evidence
    } else {
        throw "Production project '$name' has no NativeAOT classification."
    }

    [ordered]@{
        name = $name
        category = [string] $component.category
        classification = $classification
        nativeAotValidated = $nativeValidated
        evidence = $evidence
    }
}

$matrix = [ordered]@{
    schemaVersion = 1
    format = 'officeimo.nativeaot-project-matrix'
    summary = [ordered]@{
        productionProjectCount = $productionNames.Count
        nativeAotValidatedProjectCount = @($components | Where-Object nativeAotValidated).Count
        fullyRootedLibraryCount = $fullyRootedLibraries.Count
        boundedWorkflowLibraryCount = $boundedLibraries.Count
        nativeExecutableCount = $nativeTools.Count
        managedCrossPlatformProjectCount = @($managedOnly | Where-Object classification -EQ 'managed-cross-platform').Count
        managedWindowsProjectCount = @($managedOnly | Where-Object classification -EQ 'managed-windows').Count
    }
    definitions = [ordered]@{
        nativeFullSurface = 'The production library is retained as a complete assembly in the NativeAOT compile graph.'
        nativeBoundedWorkflow = 'A customer-facing workflow publishes and runs natively, but the complete optional third-party dependency surface is not claimed.'
        nativeExecutable = 'The production CLI publishes as a native executable and starts successfully.'
        managedCrossPlatform = 'The package is validated in its supported cross-platform managed deployment model rather than advertised for NativeAOT.'
        managedWindows = 'The package is validated in its supported managed Windows deployment model rather than advertised for NativeAOT.'
    }
    components = @($components)
}

if (-not [string]::IsNullOrWhiteSpace($JsonOutputPath)) {
    $resolvedOutputPath = [System.IO.Path]::GetFullPath($JsonOutputPath)
    New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutputPath) -Force | Out-Null
    $json = ($matrix | ConvertTo-Json -Depth 8).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($resolvedOutputPath, $json, [System.Text.UTF8Encoding]::new($false))
}

[pscustomobject]@{
    ProductionProjectCount = $matrix.summary.productionProjectCount
    NativeAotValidatedProjectCount = $matrix.summary.nativeAotValidatedProjectCount
    FullyRootedLibraryCount = $matrix.summary.fullyRootedLibraryCount
    BoundedWorkflowLibraryCount = $matrix.summary.boundedWorkflowLibraryCount
    NativeExecutableCount = $matrix.summary.nativeExecutableCount
    ManagedCrossPlatformProjectCount = $matrix.summary.managedCrossPlatformProjectCount
    ManagedWindowsProjectCount = $matrix.summary.managedWindowsProjectCount
    Status = 'passed'
}
