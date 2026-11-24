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

$examplesRoot = 'https://github.com/EvotecIT/OfficeIMO/tree/master/PowerShell/PSWriteOffice/Examples'

# Each family owns the commands whose noun starts with one of its prefixes, plus
# any command it lists explicitly. The longest prefix wins, so OfficeDocumentProvenance
# is not Reader, and a prefix ends at a word boundary, so OfficeAccessibility is not
# Access. Every exported command must land in exactly one family, or no catalog is written.
# 'docs' names a PSWriteOffice guide, or a site path for guides that live elsewhere;
# 'guides' adds more guides when one page does not cover the whole family, and
# 'docsTitle' then names what the main guide covers.
$familyDefinitions = @(
    [ordered]@{
        id = 'word'; title = 'Word'; description = 'Create, read and update DOC and DOCX documents: sections, paragraphs, lists, tables, headers and footers, fields, content controls, charts, review and revisions, mail merge, and conversion settings.'
        docs = 'word'; examples = "$examplesRoot/Word"
        samples = @('New-OfficeWord', 'Add-OfficeWordTable', 'Get-OfficeWordReview', 'Invoke-OfficeWordMailMerge', 'Convert-OfficeDocument')
        nouns = @('OfficeWord')
    }
    [ordered]@{
        id = 'excel'; title = 'Excel'; description = 'Build, read and update XLS and XLSX workbooks: sheets, tables, charts, pivots, slicers, validation, comments, templates, page setup, import and export, and conversion settings.'
        docs = 'excel'; examples = "$examplesRoot/Excel"
        samples = @('New-OfficeExcel', 'Add-OfficeExcelTable', 'Add-OfficeExcelPivotTable', 'Import-OfficeExcel', 'Export-OfficeExcel')
        nouns = @('OfficeExcel')
    }
    [ordered]@{
        id = 'powerpoint'; title = 'PowerPoint'; description = 'Compose, inspect and update presentations: slides, sections, layouts, placeholders, shapes, charts, tables, notes and themes, with slide import and conversion settings.'
        docs = 'powerpoint'; examples = "$examplesRoot/PowerPoint"
        samples = @('New-OfficePowerPoint', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointChart', 'Get-OfficePowerPointSlide')
        nouns = @('OfficePowerPoint')
    }
    [ordered]@{
        id = 'pdf'; title = 'PDF'; description = 'Author PDFs, read text, fonts, images and forms, edit pages, fill forms, annotate, sign, redact, add an OCR text layer, optimize, and check compliance.'
        docs = 'pdf'; examples = "$examplesRoot/Pdf"
        samples = @('New-OfficePdf', 'Get-OfficePdfFormField', 'Set-OfficePdfSignature', 'Remove-OfficePdfContent')
        nouns = @('OfficePdf', 'OfficeTextPdf')
    }
    [ordered]@{
        id = 'conversion'; title = 'Cross-format operations'; description = 'Convert, compare, merge, split, validate, repair, protect and render documents of any supported format, inspect what a file contains, and list what this installation can read and write.'
        docs = 'automation-patterns'; examples = "$examplesRoot/Workflows"
        samples = @('Convert-OfficeDocument', 'Get-OfficeDocumentInfo', 'Merge-OfficeDocument', 'Export-OfficeDocumentImage')
        nouns = @('OfficeConversion', 'OfficeVisual')
        commands = @('Convert-OfficeDocument', 'Compare-OfficeDocument', 'Get-OfficeDocumentCapability', 'Get-OfficeDocumentInfo',
            'Merge-OfficeDocument', 'Split-OfficeDocument', 'Test-OfficeDocument', 'Repair-OfficeDocument', 'Protect-OfficeDocument',
            'Unprotect-OfficeDocument', 'ConvertTo-OfficeDocumentImage', 'Export-OfficeDocumentImage')
    }
    [ordered]@{
        id = 'reader'; title = 'Reader, extraction, and OCR'; description = 'Read any supported document into one normalized model, search folders, split content into chunks and hierarchies, extract media, tables and visuals, and configure OCR.'
        docs = 'reader'; examples = "$examplesRoot/Reader"
        samples = @('Read-OfficeDocument', 'Search-OfficeDocument', 'New-OfficeDocumentReader', 'Export-OfficeDocumentAsset')
        nouns = @('OfficeDocument', 'OfficeOcr')
    }
    [ordered]@{
        id = 'before-sharing'; title = 'Before sharing'; description = 'Find and remove comments, authors, hidden content and image metadata, check Content Credentials and AI labels, and check Word, Excel and PowerPoint accessibility before a file leaves your team.'
        docs = 'before-sharing'; examples = "$examplesRoot/Workflows"
        samples = @('Get-OfficeHiddenData', 'Remove-OfficeHiddenData', 'Get-OfficeDocumentProvenance', 'Test-OfficeAccessibility')
        nouns = @('OfficeHiddenData', 'OfficeDocumentProvenance', 'OfficeAccessibility')
    }
    [ordered]@{
        id = 'printing'; title = 'Printing'; description = 'Print any document OfficeIMO can turn into PDF on a named printer queue or IPP printer, read printer capabilities, and follow or cancel jobs. Requires PowerShell 7.4 or newer.'
        docs = '/docs/pdf/operations/'; examples = "$examplesRoot/Workflows"
        samples = @('Get-OfficePrinter', 'Send-OfficePrinter', 'Get-OfficePrintJob')
        nouns = @('OfficePrinter', 'OfficePrintJob')
    }
    [ordered]@{
        id = 'notebooks'; title = 'Notebooks'; description = 'Open, edit and export Jupyter notebooks and saved results.'
        docs = 'automation-patterns'; examples = $examplesRoot
        samples = @('Open-OfficeNotebook', 'Get-OfficeNotebookCell', 'Export-OfficeNotebook')
        nouns = @('OfficeNotebook')
    }
    [ordered]@{
        id = 'visio'; title = 'Visio'; description = 'Create and update VSDX diagrams with shapes, connectors, containers and pages, use built-in or imported stencils, and turn visual definitions into editable diagrams.'
        docs = 'visio'; examples = "$examplesRoot/Visio"
        samples = @('New-OfficeVisio', 'Add-OfficeVisioStencilShape', 'Add-OfficeVisioConnector', 'Find-OfficeVisioStencil')
        nouns = @('OfficeVisio')
    }
    [ordered]@{
        id = 'project'; title = 'Project'; description = 'Create and schedule tasks, resources, assignments and dependencies, read existing plans, exchange mapped tables, export reports, and save native or XML projects.'
        docs = 'project-access-epub'; examples = "$examplesRoot/Workflows"
        samples = @('New-OfficeProject', 'Add-OfficeProjectTask', 'Invoke-OfficeProjectSchedule', 'Save-OfficeProject')
        nouns = @('OfficeProject')
    }
    [ordered]@{
        id = 'access'; title = 'Access'; description = 'Read native database schemas and rows, and create table databases from PowerShell objects without ACE or COM.'
        docs = 'project-access-epub'; examples = "$examplesRoot/Workflows"
        samples = @('New-OfficeAccess', 'Get-OfficeAccessRow', 'Add-OfficeAccessRow', 'Save-OfficeAccess')
        nouns = @('OfficeAccess')
    }
    [ordered]@{
        id = 'markdown'; title = 'Markdown'; description = 'Compose Markdown from typed blocks, parse documents into headings, tables, front matter and an object tree, and set how Markdown converts to Word, HTML, PDF, PowerPoint and RTF.'
        docs = 'open-text-formats'; examples = "$examplesRoot/Markdown"
        samples = @('New-OfficeMarkdown', 'Add-OfficeMarkdownTable', 'Get-OfficeMarkdownHeading', 'ConvertTo-OfficeMarkdown')
        nouns = @('OfficeMarkdown')
    }
    [ordered]@{
        id = 'html'; title = 'HTML'; description = 'Set how HTML is parsed, trusted and laid out when it converts to Word, PDF, Markdown or RTF, or when it is rendered to images.'
        docs = 'open-text-formats'; examples = $examplesRoot
        samples = @('New-OfficeHtmlWordOptions', 'New-OfficeHtmlPdfOptions', 'New-OfficeHtmlRenderOptions')
        nouns = @('OfficeHtml')
    }
    [ordered]@{
        id = 'rtf'; title = 'RTF'; description = 'Create, open, edit and save Rich Text Format documents, and set how RTF converts to HTML, Markdown and PDF.'
        docs = 'open-text-formats'; examples = "$examplesRoot/Rtf"
        samples = @('New-OfficeRtf', 'Open-OfficeRtf', 'Update-OfficeRtfText', 'Save-OfficeRtf')
        nouns = @('OfficeRtf')
    }
    [ordered]@{
        id = 'open-document'; title = 'OpenDocument'; description = 'Create, open, edit and save ODT text documents, ODS spreadsheets and ODP presentations.'
        docs = 'open-text-formats'; examples = $examplesRoot
        samples = @('New-OfficeOpenDocument', 'Add-OfficeOpenDocumentParagraph', 'Set-OfficeOpenDocumentCell', 'Save-OfficeOpenDocument')
        nouns = @('OfficeOpenDocument')
    }
    [ordered]@{
        id = 'email'; title = 'Email'; description = 'Read and write EML, EMLX, MSG and TNEF messages and mbox mailboxes, with reader and writer settings and fidelity diagnostics.'
        docs = 'open-text-formats'; examples = $examplesRoot
        samples = @('Open-OfficeEmail', 'Save-OfficeEmail', 'Open-OfficeEmailMailbox', 'Save-OfficeEmailMailbox')
        nouns = @('OfficeEmail')
    }
    [ordered]@{
        id = 'epub'; title = 'EPUB'; description = 'Create EPUB publications from chapters, resources and metadata, read existing chapters, and save the package.'
        docs = 'project-access-epub'; examples = "$examplesRoot/Workflows"
        samples = @('New-OfficeEpub', 'Add-OfficeEpubChapter', 'Get-OfficeEpubChapter', 'Save-OfficeEpub')
        nouns = @('OfficeEpub')
    }
    [ordered]@{
        id = 'asciidoc-latex'; title = 'AsciiDoc and LaTeX'; description = 'Parse AsciiDoc and LaTeX files or strings into the shared document model, then save them or convert them with the cross-format operations.'
        docs = 'open-text-formats'; examples = $examplesRoot
        samples = @('Open-OfficeAsciiDoc', 'Save-OfficeAsciiDoc', 'Open-OfficeLatex', 'Save-OfficeLatex')
        nouns = @('OfficeAsciiDoc', 'OfficeLatex')
    }
    [ordered]@{
        id = 'onenote-iwork'; title = 'OneNote and iWork'; description = 'Read offline OneNote sections and notebooks, and Apple Pages, Numbers and Keynote packages, without OneNote or iWork installed.'
        docs = '/docs/onenote/'; examples = $examplesRoot
        docsTitle = 'OneNote'; guides = @([ordered]@{ title = 'iWork'; url = '/docs/iwork/' })
        samples = @('Open-OfficeOneNote', 'Open-OfficeIWork')
        nouns = @('OfficeOneNote', 'OfficeIWork')
    }
    [ordered]@{
        id = 'csv'; title = 'CSV'; description = 'Import, export and convert delimited data with explicit quoting, encoding and formula-injection rules.'
        docs = 'config-data-formats'; examples = "$examplesRoot/Csv"
        samples = @('Import-OfficeCsv', 'Export-OfficeCsv', 'ConvertTo-OfficeCsv', 'ConvertFrom-OfficeCsv')
        nouns = @('OfficeCsv')
    }
    [ordered]@{
        id = 'yaml'; title = 'YAML'; description = 'Edit YAML files without losing comments or layout, merge overlays, validate, and convert between YAML and objects.'
        docs = 'config-data-formats'; examples = "$examplesRoot/Yaml"
        samples = @('Open-OfficeYaml', 'Set-OfficeYamlNode', 'ConvertFrom-OfficeYaml', 'Test-OfficeYaml')
        nouns = @('OfficeYaml')
    }
    [ordered]@{
        id = 'toml'; title = 'TOML'; description = 'Edit TOML files without losing comments or layout, and read and write TOML configuration and lock files as PowerShell objects.'
        docs = 'config-data-formats'; examples = "$examplesRoot/Toml"
        samples = @('Open-OfficeToml', 'Set-OfficeTomlNode', 'Import-OfficeToml', 'Export-OfficeToml')
        nouns = @('OfficeToml')
    }
    [ordered]@{
        id = 'confluence'; title = 'Confluence Cloud'; description = 'Plan and publish pages, read and remove them, preserve managed sections, and list, upload or download attachments.'
        docs = 'confluence'; examples = "$examplesRoot/Confluence"
        samples = @('New-OfficeConfluenceSession', 'Publish-OfficeConfluencePage', 'Set-OfficeConfluenceManagedSection', 'Send-OfficeConfluenceAttachment')
        nouns = @('OfficeConfluence')
    }
    [ordered]@{
        id = 'google-workspace'; title = 'Google Docs and Sheets'; description = 'Sign in with a service account, then plan or export Word documents to Google Docs and Excel workbooks to Google Sheets.'
        docs = '/docs/google-workspace/'; examples = $examplesRoot
        samples = @('New-OfficeGoogleWorkspaceSession', 'Export-OfficeWordGoogleDocument', 'Export-OfficeExcelGoogleSpreadsheet')
        nouns = @('OfficeGoogleWorkspace')
        commands = @('Export-OfficeWordGoogleDocument', 'New-OfficeWordGoogleDocsOptions',
            'Export-OfficeExcelGoogleSpreadsheet', 'New-OfficeExcelGoogleSheetsOptions')
    }
)

# Groups say what a command does inside its family. The verb decides, except that
# New-*Options and New-*Filter commands only build settings for other commands.
# A verb that is in no group stops the build, so it is placed deliberately.
$groupDefinitions = @(
    [ordered]@{ id = 'create'; title = 'Create, open and save'; verbs = @('New', 'Open', 'Save', 'Close') }
    [ordered]@{ id = 'add'; title = 'Add content'; verbs = @('Add') }
    [ordered]@{ id = 'read'; title = 'Read and inspect'; verbs = @('Get', 'Find', 'Search', 'Read', 'Import', 'ConvertFrom') }
    [ordered]@{ id = 'edit'; title = 'Edit and organize'; verbs = @('Set', 'Update', 'Remove', 'Clear', 'Copy', 'Move', 'Rename', 'Edit', 'Invoke', 'Merge', 'Split', 'Resolve', 'Optimize', 'Stop') }
    [ordered]@{ id = 'convert'; title = 'Convert, export and publish'; verbs = @('Convert', 'ConvertTo', 'Export', 'Publish', 'Send') }
    [ordered]@{ id = 'check'; title = 'Validate, protect and compare'; verbs = @('Test', 'Compare', 'Repair', 'Protect', 'Unprotect') }
    [ordered]@{ id = 'options'; title = 'Settings for other commands'; verbs = @() }
)

function Get-CommandGroupId {
    param([Parameter(Mandatory)][string] $Name)

    $verb, $noun = $Name -split '-', 2
    if ($verb -ceq 'New' -and $noun -cmatch '(Options|Filter)$') { return 'options' }
    foreach ($group in $groupDefinitions) {
        if ($group.verbs -ccontains $verb) { return $group.id }
    }
    throw "The command '$Name' uses the verb '$verb', which has no documentation group. Add the verb to a group in Export-WebsiteDocumentationCatalog.ps1."
}

function Test-NounPrefix {
    param([string] $Noun, [string] $Prefix)

    $Noun.StartsWith($Prefix, [StringComparison]::Ordinal) -and
        ($Noun.Length -eq $Prefix.Length -or [char]::IsUpper($Noun[$Prefix.Length]))
}

$familyIdByCommand = @{}
foreach ($definition in $familyDefinitions) {
    foreach ($command in @($definition.commands | Where-Object { $_ })) {
        if ($commands -cnotcontains $command) {
            throw "Documentation family '$($definition.id)' lists '$command', which PSWriteOffice does not export."
        }
        if ($familyIdByCommand.ContainsKey($command)) {
            throw "'$command' is listed by both '$($familyIdByCommand[$command])' and '$($definition.id)'."
        }
        $familyIdByCommand[$command] = $definition.id
    }
}
$unassigned = [System.Collections.Generic.List[string]]::new()
foreach ($command in $commands) {
    if ($familyIdByCommand.ContainsKey($command)) { continue }
    $noun = ($command -split '-', 2)[1]
    $owner = $null
    $ownerPrefixLength = 0
    foreach ($definition in $familyDefinitions) {
        foreach ($prefix in @($definition.nouns)) {
            if ($prefix.Length -gt $ownerPrefixLength -and (Test-NounPrefix -Noun $noun -Prefix $prefix)) {
                $owner = $definition.id
                $ownerPrefixLength = $prefix.Length
            }
        }
    }
    if ($owner) { $familyIdByCommand[$command] = $owner } else { $unassigned.Add($command) }
}
if ($unassigned.Count -gt 0) {
    throw "Commands are missing a documentation family: $($unassigned -join ', '). Give a family a matching noun prefix or list the command explicitly in Export-WebsiteDocumentationCatalog.ps1."
}

$families = foreach ($definition in $familyDefinitions) {
    $familyCommands = @($commands | Where-Object { $familyIdByCommand[$_] -ceq $definition.id })
    if ($familyCommands.Count -eq 0) {
        throw "Documentation family '$($definition.id)' has no exported commands."
    }

    $missingSamples = @($definition.samples | Where-Object { $commands -cnotcontains $_ })
    if ($missingSamples.Count -gt 0) {
        throw "Documentation catalog for '$($definition.id)' references commands not exported by PSWriteOffice: $($missingSamples -join ', ')"
    }

    $groups = foreach ($group in $groupDefinitions) {
        $groupCommands = @($familyCommands | Where-Object { (Get-CommandGroupId -Name $_) -ceq $group.id })
        if ($groupCommands.Count -gt 0) {
            [ordered]@{ id = $group.id; title = $group.title; commands = $groupCommands }
        }
    }

    [ordered]@{
        id = $definition.id
        title = $definition.title
        description = $definition.description
        commandCount = $familyCommands.Count
        docsUrl = if ($definition.docs.StartsWith('/')) { $definition.docs } else { "/docs/pswriteoffice/$($definition.docs)/" }
        docsTitle = if ($definition.docsTitle) { $definition.docsTitle } else { $definition.title }
        apiUrl = '/api/powershell/'
        examplesUrl = $definition.examples
        featuredCommands = @($definition.samples)
        guides = @($definition.guides | Where-Object { $_ })
        groups = @($groups)
    }
}

$catalog = [ordered]@{
    schemaVersion = 2
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
