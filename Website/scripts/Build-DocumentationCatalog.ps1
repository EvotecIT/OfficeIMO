param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path,
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $OutputPath = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $SiteRoot 'data\documentation_catalog.json'
}

function Get-ProjectValue {
    param([xml] $Project, [string] $Name)

    $values = @(@($Project.Project.PropertyGroup.$Name) |
        ForEach-Object { [string] $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($values.Count -gt 0) { return $values[-1].Trim() }
    return ''
}

function Get-ComponentCategory {
    param([string] $Name)

    switch -Regex ($Name) {
        '^OfficeIMO\.Reader' { return 'Extraction and ingestion' }
        '^OfficeIMO\.(Word|Excel|PowerPoint)' { return 'Office documents' }
        '^OfficeIMO\.(Pdf|Html|Markdown|Rtf|AsciiDoc|Latex)' { return 'Publishing and conversion' }
        '^OfficeIMO\.(Email|OneNote|OpenDocument|Epub|CSV|Visio)' { return 'Formats and interoperability' }
        '^OfficeIMO\.GoogleWorkspace|Google(Docs|Sheets|Slides)$' { return 'Google Workspace' }
        '^OfficeIMO\.(Drawing|Security|Zip|Markup|Adf|Confluence)' { return 'Foundations and integrations' }
        '^OfficeIMO\.MarkdownRenderer' { return 'Rendering surfaces' }
        default { return 'Specialized components' }
    }
}

function Get-DocumentationUrl {
    param([string] $Name)

    switch -Regex ($Name) {
        '^OfficeIMO\.Word' { return '/docs/word/' }
        '^OfficeIMO\.Excel' { return '/docs/excel/' }
        '^OfficeIMO\.PowerPoint' { return '/docs/powerpoint/' }
        '^OfficeIMO\.Pdf' { return '/docs/pdf/' }
        '^OfficeIMO\.Email' { return '/docs/email/' }
        '^OfficeIMO\.OneNote' { return '/docs/onenote/' }
        '^OfficeIMO\.Html' { return '/docs/html/' }
        '^OfficeIMO\.OpenDocument' { return '/docs/open-document/' }
        '^OfficeIMO\.Markdown' { return '/docs/markdown/' }
        '^OfficeIMO\.CSV' { return '/docs/csv/' }
        '^OfficeIMO\.Visio' { return '/docs/visio/' }
        '^OfficeIMO\.Reader' { return '/docs/reader/' }
        '^OfficeIMO\.GoogleWorkspace|Google(Docs|Sheets|Slides)$' { return '/docs/google-workspace/' }
        default { return '/docs/capabilities/packages/' }
    }
}

$allProjects = @(Get-ChildItem -LiteralPath $RepositoryRoot -Recurse -File -Filter '*.csproj' |
    Where-Object { $_.FullName -notmatch '\\(?:bin|obj|Website\\projects)\\' })

$testProjects = @($allProjects | Where-Object { $_.BaseName -match '(?:^|\.)(?:Tests|VerifyTests)(?:\.|$)' })
$benchmarkProjects = @($allProjects | Where-Object { $_.BaseName -match '(?:^|\.)Benchmarks(?:\.|$)' })
$validationProjects = @($allProjects | Where-Object {
    $_.FullName -match '\\Build\\' -or
    $_.BaseName -match '(?:^|\.)AotSmoke$' -or
    $_.BaseName -in @('OfficeIMO.Examples', 'OfficeIMO.MarkdownRenderer.SamplePlugin') -or
    $_.FullName -match '\\Website\\Apps\\'
})
$productionProjects = @($allProjects | Where-Object {
    $_ -notin $testProjects -and
    $_ -notin $benchmarkProjects -and
    $_ -notin $validationProjects
})

$referenceCounts = @{}
foreach ($testProject in $testProjects) {
    [xml] $testXml = Get-Content -LiteralPath $testProject.FullName -Raw
    foreach ($reference in @($testXml.Project.ItemGroup.ProjectReference)) {
        $include = [string] $reference.Include
        if ([string]::IsNullOrWhiteSpace($include)) { continue }
        $referencedPath = [System.IO.Path]::GetFullPath((Join-Path $testProject.DirectoryName $include))
        if (-not $referenceCounts.ContainsKey($referencedPath)) { $referenceCounts[$referencedPath] = 0 }
        $referenceCounts[$referencedPath]++
    }
}

$pipelinePath = Join-Path $SiteRoot 'pipeline.json'
$pipeline = Get-Content -LiteralPath $pipelinePath -Raw | ConvertFrom-Json
$apiRoutes = @{}
foreach ($step in @($pipeline.steps | Where-Object task -EQ 'apidocs')) {
    $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension([string] $step.assembly)
    if (-not [string]::IsNullOrWhiteSpace($assemblyName)) {
        $apiRoutes[$assemblyName] = ([string] $step.baseUrl).TrimEnd('/') + '/'
    }
}

$components = foreach ($projectFile in ($productionProjects | Sort-Object BaseName)) {
    [xml] $project = Get-Content -LiteralPath $projectFile.FullName -Raw
    $name = $projectFile.BaseName
    $packageId = Get-ProjectValue -Project $project -Name 'PackageId'
    if ([string]::IsNullOrWhiteSpace($packageId)) { $packageId = $name }
    $description = Get-ProjectValue -Project $project -Name 'Description'
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = "$name is a focused component in the OfficeIMO managed document platform."
    }
    $isPackable = (Get-ProjectValue -Project $project -Name 'IsPackable') -ne 'false'
    $outputType = Get-ProjectValue -Project $project -Name 'OutputType'
    $projectPath = [System.IO.Path]::GetRelativePath($RepositoryRoot, $projectFile.FullName).Replace('\', '/')
    $sourceCount = @(Get-ChildItem -LiteralPath $projectFile.DirectoryName -Recurse -File -Filter '*.cs' |
        Where-Object { $_.FullName -notmatch '\\(?:bin|obj)\\' }).Count
    $resolvedProjectPath = [System.IO.Path]::GetFullPath($projectFile.FullName)

    [ordered]@{
        name = $name
        packageId = $packageId
        category = Get-ComponentCategory -Name $name
        description = $description
        kind = if ($outputType -eq 'Exe') { 'tool' } elseif ($isPackable) { 'package' } else { 'library' }
        projectPath = $projectPath
        sourceFileCount = $sourceCount
        referencingTestProjectCount = if ($referenceCounts.ContainsKey($resolvedProjectPath)) { $referenceCounts[$resolvedProjectPath] } else { 0 }
        docsUrl = Get-DocumentationUrl -Name $name
        apiUrl = if ($apiRoutes.ContainsKey($name)) { $apiRoutes[$name] } else { $null }
        packageUrl = if ($isPackable) { "https://www.nuget.org/packages/$packageId" } else { $null }
    }
}

$categories = @($components | Group-Object { $_['category'] } | Sort-Object Name | ForEach-Object {
    [ordered]@{ name = $_.Name; componentCount = $_.Count }
})
$powerShellApiAvailable = Test-Path -LiteralPath (Join-Path $SiteRoot 'data\apidocs\powershell\command-metadata.json') -PathType Leaf
$apiReferenceCount = $apiRoutes.Count + $(if ($powerShellApiAvailable) { 1 } else { 0 })

$capabilitiesRoot = Join-Path $SiteRoot 'content\docs\capabilities'
$packagesRoot = Join-Path $capabilitiesRoot 'packages'
New-Item -ItemType Directory -Path $packagesRoot -Force | Out-Null

$overview = [System.Text.StringBuilder]::new()
[void] $overview.AppendLine('---')
[void] $overview.AppendLine('title: "OfficeIMO Capability Catalog"')
[void] $overview.AppendLine('description: "Navigate the complete managed document platform by workflow, component family, generated API, and validation evidence."')
[void] $overview.AppendLine('layout: docs')
[void] $overview.AppendLine('---')
[void] $overview.AppendLine()
[void] $overview.AppendLine("OfficeIMO is a modular document platform, not a single basic DOCX helper. The repository currently contains **$($components.Count) production libraries, adapters, renderers, and tools**, backed by **$($testProjects.Count) test projects**, **$($benchmarkProjects.Count) benchmark projects**, and dedicated validation applications.")
[void] $overview.AppendLine()
[void] $overview.AppendLine('## Choose the right depth')
[void] $overview.AppendLine()
[void] $overview.AppendLine('| Need | Start here |')
[void] $overview.AppendLine('|---|---|')
[void] $overview.AppendLine('| Build or edit a document | Use the Word, Excel, PowerPoint, PDF, email, OneNote, OpenDocument, Markdown, CSV, or Visio guide. |')
[void] $overview.AppendLine('| Move content between formats | Open the conversion map to choose the source package, destination adapter, and expected loss policy. |')
[void] $overview.AppendLine('| Normalize mixed documents | Start with Reader and add only the format adapters your application needs. |')
[void] $overview.AppendLine('| Automate from scripts | Use PSWriteOffice and its manifest-derived 464-command catalog. |')
[void] $overview.AppendLine('| Inspect exact members | Move from a conceptual guide into one of the generated API references. |')
[void] $overview.AppendLine('| Evaluate deployment constraints | Use the validation and AOT pages for executable evidence and known boundaries. |')
[void] $overview.AppendLine()
[void] $overview.AppendLine('## Component families')
[void] $overview.AppendLine()
foreach ($category in $categories) {
    [void] $overview.AppendLine("- **$($category.name):** $($category.componentCount) focused components")
}
[void] $overview.AppendLine()
[void] $overview.AppendLine('The [complete package and component index](./packages/) is generated from project metadata, so descriptions, source links, API availability, and test-reference counts stay aligned with the repository.')
[void] $overview.AppendLine()
[void] $overview.AppendLine('## Evidence, not blanket promises')
[void] $overview.AppendLine()
[void] $overview.AppendLine('Support is stated at the scenario level. A package can be cross-platform while a renderer, OCR provider, browser runtime, authentication provider, or NativeAOT path has additional constraints. The [validation guide](./validation/) separates source compatibility, automated tests, published artifacts, and executed deployment paths.')
Set-Content -LiteralPath (Join-Path $capabilitiesRoot 'index.md') -Value $overview.ToString() -Encoding utf8 -NoNewline

$packagePage = [System.Text.StringBuilder]::new()
[void] $packagePage.AppendLine('---')
[void] $packagePage.AppendLine('title: "Packages and Components"')
[void] $packagePage.AppendLine('description: "A source-derived index of every production OfficeIMO library, adapter, renderer, and tool."')
[void] $packagePage.AppendLine('layout: docs')
[void] $packagePage.AppendLine('---')
[void] $packagePage.AppendLine()
[void] $packagePage.AppendLine('This page is generated from the repository project files. It includes focused engines, conversion adapters, Reader providers, rendering surfaces, command-line tools, and integration packages that are easy to miss when only the top-level product pages are visible.')
foreach ($categoryName in @($categories.name)) {
    [void] $packagePage.AppendLine()
    [void] $packagePage.AppendLine("## $categoryName")
    [void] $packagePage.AppendLine()
    [void] $packagePage.AppendLine('| Component | What it owns | Evidence and next steps |')
    [void] $packagePage.AppendLine('|---|---|---|')
    foreach ($component in @($components | Where-Object { $_['category'] -eq $categoryName } | Sort-Object { $_['name'] })) {
        $name = [string] $component['name']
        $description = ([string] $component['description']).Replace('|', '\|')
        $links = [System.Collections.Generic.List[string]]::new()
        $componentName = [string] $component['name']
        $links.Add("[$componentName guide]($($component['docsUrl']))")
        if ($component['apiUrl']) { $links.Add("[$componentName API]($($component['apiUrl']))") }
        if ($component['packageUrl']) { $links.Add("[$componentName package]($($component['packageUrl']))") }
        $links.Add("$($component['sourceFileCount']) source files")
        if ([int] $component['referencingTestProjectCount'] -gt 0) {
            $links.Add("$($component['referencingTestProjectCount']) referencing test project(s)")
        }
        [void] $packagePage.AppendLine("| ``$name`` | $description | $($links -join ' · ') |")
    }
}
[void] $packagePage.AppendLine()
[void] $packagePage.AppendLine('Project metadata proves that a component exists and identifies its intended ownership. Use the linked conceptual guide, generated API, examples, and tests to validate the exact scenario you plan to ship.')
Set-Content -LiteralPath (Join-Path $packagesRoot 'index.md') -Value $packagePage.ToString() -Encoding utf8 -NoNewline

$docs = @(Get-ChildItem -LiteralPath (Join-Path $SiteRoot 'content\docs') -Recurse -File -Filter '*.md')
$wordCount = 0
foreach ($doc in $docs) {
    $text = Get-Content -LiteralPath $doc.FullName -Raw
    $wordCount += @([regex]::Matches($text, '\b[\p{L}\p{N}][\p{L}\p{N}\-\.]*\b')).Count
}

$catalog = [ordered]@{
    schemaVersion = 1
    format = 'officeimo.documentation-catalog'
    repository = [ordered]@{
        projectCount = $allProjects.Count
        productionComponentCount = $components.Count
        testProjectCount = $testProjects.Count
        benchmarkProjectCount = $benchmarkProjects.Count
        validationProjectCount = $validationProjects.Count
        apiReferenceCount = $apiReferenceCount
        conceptualPageCount = $docs.Count
        conceptualWordCount = $wordCount
    }
    categories = $categories
    components = @($components)
}

$parent = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $parent -Force | Out-Null
$catalog | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8

[PSCustomObject]@{
    OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path
    ProjectCount = $allProjects.Count
    ProductionComponentCount = $components.Count
    TestProjectCount = $testProjects.Count
    ApiReferenceCount = $apiReferenceCount
    ConceptualPageCount = $docs.Count
}
