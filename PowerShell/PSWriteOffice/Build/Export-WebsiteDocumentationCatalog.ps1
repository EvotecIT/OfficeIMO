param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $OutputPath = '',
    [string] $ManifestPath = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $RepositoryRoot 'WebsiteArtifacts\documentation\command-catalog.json'
}

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $RepositoryRoot 'PSWriteOffice.psd1'
}
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "PSWriteOffice manifest was not found at '$ManifestPath'."
}

$manifest = Import-PowerShellDataFile -LiteralPath $ManifestPath
$commands = @($manifest.CmdletsToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
    Sort-Object -Unique
$aliases = @($manifest.AliasesToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
    Sort-Object -Unique

$familyDefinitions = @(
    [ordered]@{
        id = 'project'; title = 'Project'; description = 'Create schedules, inspect tasks and resources, exchange mapped tables, and save qualified native or XML projects.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Workflows'
        samples = @('New-OfficeProject', 'Invoke-OfficeProjectSchedule', 'Save-OfficeProject', 'Import-OfficeProjectTable')
        match = { param($name) $name -match 'OfficeProject' }
    }
    [ordered]@{
        id = 'access'; title = 'Access'; description = 'Read native database schemas and rows, and create qualified table databases without ACE or COM.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Workflows'
        samples = @('New-OfficeAccess', 'Get-OfficeAccessRow', 'Add-OfficeAccessRow', 'Save-OfficeAccess')
        match = { param($name) $name -match 'OfficeAccess' }
    }
    [ordered]@{
        id = 'epub'; title = 'EPUB'; description = 'Read, create, and save EPUB publications from chapters, resources, and metadata.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Workflows'
        samples = @('New-OfficeEpub', 'Add-OfficeEpubChapter', 'Get-OfficeEpubChapter', 'Save-OfficeEpub')
        match = { param($name) $name -match 'OfficeEpub' }
    }
    [ordered]@{
        id = 'conversion'; title = 'Conversion and capabilities'; description = 'Discover executable routes and protected-content support, and run bounded single-file, selected-file, and resumable directory jobs.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Workflows'
        samples = @('Get-OfficeDocumentCapability', 'New-OfficeConversionOptions', 'Convert-OfficeDocument')
        match = { param($name) $name -match 'OfficeConversion' -or $name -in 'Convert-OfficeDocument', 'ConvertTo-OfficeDocument', 'Get-OfficeDocumentCapability' }
    }
    [ordered]@{
        id = 'word'; title = 'Word'; description = 'Create, inspect, update, review, merge, protect, and convert DOC and DOCX documents.'
        docs = 'word'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Word'
        samples = @('New-OfficeWord', 'Add-OfficeWordTable', 'Get-OfficeWordReview', 'Invoke-OfficeWordMailMerge', 'ConvertTo-OfficeWordDocument')
        match = { param($name) $name -match 'OfficeWord' }
    }
    [ordered]@{
        id = 'excel'; title = 'Excel'; description = 'Build, read, convert, validate, repair, compare, and publish XLS and XLSX workbook workflows.'
        docs = 'excel'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Excel'
        samples = @('New-OfficeExcel', 'Add-OfficeExcelTable', 'Add-OfficeExcelPivotTable', 'Test-OfficeExcelWorkbook', 'ConvertTo-OfficeExcelWorkbook')
        match = { param($name) $name -match 'OfficeExcel' }
    }
    [ordered]@{
        id = 'powerpoint'; title = 'PowerPoint'; description = 'Compose, inspect, update, import, theme, and render repeatable presentation decks.'
        docs = 'powerpoint'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/PowerPoint'
        samples = @('New-OfficePowerPoint', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointChart', 'Get-OfficePowerPointSlide')
        match = { param($name) $name -match 'OfficePowerPoint' }
    }
    [ordered]@{
        id = 'pdf'; title = 'PDF'; description = 'Author, inspect, transform, sign, annotate, extract, preflight, and combine PDF files.'
        docs = 'pdf'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Pdf'
        samples = @('New-OfficePdf', 'Join-OfficePdf', 'Get-OfficePdfPreflight', 'Set-OfficePdfSignature')
        match = { param($name) $name -match 'OfficePdf|OfficeDocumentPdf|OfficeTextPdf|OfficePrinter' }
    }
    [ordered]@{
        id = 'reader'; title = 'Reader, extraction, and OCR'; description = 'Detect formats and extract normalized documents, chunks, tables, visuals, assets, OCR text, and ingest results.'
        docs = 'reader'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Reader'
        samples = @('New-OfficeDocumentReader', 'Read-OfficeDocument', 'Export-OfficeDocumentAsset', 'Search-OfficeDocument')
        match = { param($name) ($name -match 'OfficeDocument' -and $name -ne 'Export-OfficeDocumentPdf') -or $name -match 'OfficeReader|OfficeOcr|OfficeImageText|OfficeImageDocument|^ConvertFrom-OfficeImage$' }
    }
    [ordered]@{
        id = 'archive-migration'; title = 'Archive migration and evidence'; description = 'Inspect package safety and provenance, convert iWork and offline OneNote content, and retain fidelity evidence.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('Get-OfficePackageSecurity', 'Get-OfficeDocumentProvenance', 'Get-OfficeIWork', 'ConvertFrom-OfficeOneNote')
        match = { param($name) $name -match 'OfficeIWork|OfficeOneNote|OfficePackageSecurity|OfficeProvenance' }
    }
    [ordered]@{
        id = 'confluence'; title = 'Confluence Cloud'; description = 'Plan and publish pages, preserve managed sections, and transfer attachments through OfficeIMO.Confluence.'
        docs = 'confluence'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Confluence'
        samples = @('New-OfficeConfluenceSession', 'Publish-OfficeConfluencePage', 'Set-OfficeConfluenceManagedSection', 'Send-OfficeConfluenceAttachment')
        match = { param($name) $name -match 'OfficeConfluence' }
    }
    [ordered]@{
        id = 'visio'; title = 'Visio'; description = 'Create, inspect, arrange, and export VSDX diagrams with built-in and imported stencils.'
        docs = 'visio'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Visio'
        samples = @('New-OfficeVisio', 'Add-OfficeVisioStencilShape', 'Get-OfficeVisioInfo', 'Export-OfficeDocumentImage')
        match = { param($name) $name -match 'OfficeVisio' }
    }
    [ordered]@{
        id = 'markdown'; title = 'Markdown'; description = 'Compose typed Markdown, parse documents, and convert between Markdown and HTML or Word.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Markdown'
        samples = @('New-OfficeMarkdown', 'Add-OfficeMarkdownTable', 'ConvertTo-OfficeMarkdownHtml', 'ConvertFrom-OfficeWordMarkdown')
        match = { param($name) $name -match 'Markdown' -and $name -notmatch 'OfficeAsciiDoc|OfficeLatex' }
    }
    [ordered]@{
        id = 'rtf'; title = 'RTF'; description = 'Create, open, edit, inspect, and bridge Rich Text Format documents.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Rtf'
        samples = @('New-OfficeRtf', 'Get-OfficeRtf', 'Update-OfficeRtfText', 'ConvertTo-OfficeRtf')
        match = { param($name) $name -match 'OfficeRtf' }
    }
    [ordered]@{
        id = 'csv'; title = 'CSV'; description = 'Create, import, inspect, and export delimited data with typed options.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples/Csv'
        samples = @('ConvertTo-OfficeCsv', 'Import-OfficeCsv', 'Export-OfficeCsv', 'Get-OfficeCsv')
        match = { param($name) $name -match 'OfficeCsv' }
    }
    [ordered]@{
        id = 'open-document'; title = 'OpenDocument'; description = 'Compose, read, convert, and save ODT, ODS, and ODP workflows.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('New-OfficeOpenDocument', 'Add-OfficeOpenDocumentParagraph', 'Set-OfficeOpenDocumentCell', 'Save-OfficeOpenDocument')
        match = { param($name) $name -match 'OfficeOpenDocument' }
    }
    [ordered]@{
        id = 'email'; title = 'Email'; description = 'Read and write messages and mailbox artifacts through the managed OfficeIMO.Email engine.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('Get-OfficeEmail', 'Get-OfficeEmailMailbox', 'Save-OfficeEmail', 'Save-OfficeEmailMailbox')
        match = { param($name) $name -match 'OfficeEmail' }
    }
    [ordered]@{
        id = 'asciidoc'; title = 'AsciiDoc'; description = 'Read, create, update, and save bounded AsciiDoc workflows.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('Get-OfficeAsciiDoc', 'ConvertFrom-OfficeAsciiDocMarkdown', 'ConvertTo-OfficeAsciiDocMarkdown', 'Save-OfficeAsciiDoc')
        match = { param($name) $name -match 'OfficeAsciiDoc' }
    }
    [ordered]@{
        id = 'latex'; title = 'LaTeX'; description = 'Read, create, update, and save bounded LaTeX interoperability workflows.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('Get-OfficeLatex', 'ConvertFrom-OfficeLatexMarkdown', 'ConvertTo-OfficeLatexMarkdown', 'Save-OfficeLatex')
        match = { param($name) $name -match 'OfficeLatex' }
    }
    [ordered]@{
        id = 'html'; title = 'HTML assets'; description = 'Export images and review surfaces used by document-to-HTML workflows.'
        docs = 'open-text-formats'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('Export-OfficeHtmlImage')
        match = { param($name) $name -match 'OfficeHtml' }
    }
    [ordered]@{
        id = 'visuals'; title = 'Cross-format visuals'; description = 'Convert reusable visual artifacts for Word, Excel, PowerPoint, and PDF placement.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('ConvertTo-OfficeVisual')
        match = { param($name) $name -eq 'ConvertTo-OfficeVisual' }
    }
    [ordered]@{
        id = 'shared'; title = 'Shared authoring primitives'; description = 'Create reusable text runs shared by document DSLs.'
        docs = 'automation-patterns'; api = '/api/powershell/'; examples = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'
        samples = @('New-OfficeTextRun')
        match = { param($name) $name -eq 'New-OfficeTextRun' }
    }
)

$assigned = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$families = foreach ($definition in $familyDefinitions) {
    # Cross-format commands can mention more than one family (for example,
    # ConvertFrom-OfficeWordMarkdown). The ordered family list establishes one
    # documentation owner so totals remain trustworthy and every command has a
    # single place in the navigation.
    $familyCommands = @($commands | Where-Object {
        -not $assigned.Contains($_) -and (& $definition.match $_)
    })
    foreach ($command in $familyCommands) { [void] $assigned.Add($command) }

    $missingSamples = @($definition.samples | Where-Object { $_ -notin $commands })
    if ($missingSamples.Count -gt 0) {
        throw "Documentation catalog for '$($definition.id)' references commands not exported by PSWriteOffice: $($missingSamples -join ', ')"
    }

    [ordered]@{
        id = $definition.id
        title = $definition.title
        description = $definition.description
        commandCount = $familyCommands.Count
        docsUrl = "/docs/pswriteoffice/$($definition.docs)/"
        apiUrl = $definition.api
        examplesUrl = $definition.examples
        featuredCommands = @($definition.samples)
    }
}

$unassigned = @($commands | Where-Object { -not $assigned.Contains($_) })
if ($unassigned.Count -gt 0) {
    throw "Commands are missing a documentation family: $($unassigned -join ', ')"
}

$catalog = [ordered]@{
    schemaVersion = 1
    format = 'pswriteoffice.documentation-catalog'
    module = [ordered]@{
        name = 'PSWriteOffice'
        version = [string] $manifest.ModuleVersion
        commandCount = $commands.Count
        aliasCount = $aliases.Count
        familyCount = $families.Count
        sourceManifest = 'PSWriteOffice.psd1'
    }
    families = @($families)
}

[System.Management.Automation.ProviderInfo] $outputProvider = $null
[System.Management.Automation.PSDriveInfo] $outputDrive = $null
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutputPath, [ref] $outputProvider, [ref] $outputDrive)
if ($outputProvider.Name -ne 'FileSystem') {
    throw 'OutputPath must use the FileSystem provider.'
}
$parent = [System.IO.Path]::GetDirectoryName($resolvedOutputPath)
if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$catalogJson = $catalog | ConvertTo-Json -Depth 8 -Compress
$utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText(
    $resolvedOutputPath,
    $catalogJson + "`n",
    $utf8WithoutBom)

[PSCustomObject]@{
    OutputPath = $resolvedOutputPath
    CommandCount = $commands.Count
    AliasCount = $aliases.Count
    FamilyCount = $families.Count
}
