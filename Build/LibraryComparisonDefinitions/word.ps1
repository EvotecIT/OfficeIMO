# Word (DOCX) suites beyond the core create, report, read and replace comparisons in LibraryComparisonDefinitions.ps1.
param(
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0'
)

$project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
$platforms = @('windows', 'linux', 'macos')
# OneDocument ships net10.0 only, so its lanes and library name are declared only for that target.
$one = $Framework -eq 'net10.0'

function New-WordExpectedCases {
    param([string[]] $Methods, [string] $Parameter, [object[]] $Values)
    foreach ($value in $Values) {
        foreach ($method in $Methods) { "$method|$Parameter=$value" }
    }
}

[ordered]@{
    wordimages = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX creation with inline images'
        Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*WordCreateImageComparisonBenchmarks*'
        ComparisonId = 'word-docx-create-images'
        Suite = 'OfficeIMO.Word.CreateImages'
        # The Open XML SDK lane writes only the document and image parts, like the Minimal lanes of the paragraph
        # and report comparisons.
        MethodIdentity = @{ 'OpenXmlSdk' = 'Open XML SDK|Minimal' }
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdk'; if ($one) { 'OneDocument' }) -Parameter 'ImageCount' -Values @(10, 100))
    }
    wordtemplatefill = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX template fill and save'
        Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*WordTemplateFillComparisonBenchmarks*'
        ComparisonId = 'word-docx-template-fill'
        Suite = 'OfficeIMO.Word.TemplateFill'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdk'; if ($one) { 'OneDocument' }) -Parameter 'ItemCount' -Values @(100, 1000))
    }
    wordreadtables = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX table cell reading'
        Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*WordReadTableComparisonBenchmarks*'
        ComparisonId = 'word-docx-read-tables'
        Suite = 'OfficeIMO.Word.ReadTables'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdk'; if ($one) { 'OneDocument' }) -Parameter 'RowCount' -Values @(100, 1000))
    }
    wordevidenceparagraphs = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX paragraph creation (Open XML SDK on an OfficeIMO package shell)'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'The Open XML SDK lane appends paragraphs to a package shell OfficeIMO wrote instead of creating the package, so the lanes do not do the same creation work; word-docx-create-paragraphs is the published creation comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordParagraphOpenXmlEvidenceBenchmarks*'
        ComparisonId = 'word-docx-evidence-paragraphs'
        Suite = 'OfficeIMO.Word.EvidenceParagraphs'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'OpenXmlSdk') -Parameter 'ItemCount' -Values @(100, 1000))
    }
    wordevidencereport = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX report creation (Open XML SDK on an OfficeIMO package shell)'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'The Open XML SDK lane adds the report to a package shell OfficeIMO wrote instead of creating the package, so the lanes do not do the same creation work; word-docx-create-report is the published report comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordReportOpenXmlEvidenceBenchmarks*'
        ComparisonId = 'word-docx-evidence-report'
        Suite = 'OfficeIMO.Word.EvidenceReport'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'OpenXmlSdk') -Parameter 'RowCount' -Values @(100, 1000))
    }
    wordevidenceread = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX paragraph reading (OfficeIMO and Open XML SDK lanes)'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'Repeats the OfficeIMO and Open XML SDK lanes of word-docx-read-paragraphs over the same fixture at 100 and 1000 items; publishing it would list the same measurement twice.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordReadOpenXmlEvidenceBenchmarks*'
        ComparisonId = 'word-docx-evidence-read'
        Suite = 'OfficeIMO.Word.EvidenceRead'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'OpenXmlSdk') -Parameter 'ItemCount' -Values @(100, 1000))
    }
    wordreplacestyled = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX text replacement in a fully styled package'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*WordReplaceOpenXmlEvidenceBenchmarks*'
        ComparisonId = 'word-docx-replace-styled'
        Suite = 'OfficeIMO.Word.ReplaceStyled'
        ExpectedCases = @(New-WordExpectedCases -Methods @('OfficeIMO', 'OpenXmlSdk') -Parameter 'ItemCount' -Values @(100, 1000))
    }
    wordhtmlimport = [pscustomobject]@{
        Area = 'word'
        Title = 'HTML to DOCX conversion'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only scaling measurement; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordHtmlImportBenchmarks*'
        ComparisonId = 'word-html-import'
        Suite = 'OfficeIMO.Word.HtmlImport'
        ExpectedCases = @(
            foreach ($shape in @('Containers', 'Tables', 'Lists')) {
                foreach ($blockCount in @(100, 400)) { "HtmlToWord|Shape=$shape&BlockCount=$blockCount" }
            }
        )
    }
    wordload = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX load with images'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only load-mode measurement; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordLoadBenchmarks*'
        ComparisonId = 'word-load-images'
        Suite = 'OfficeIMO.Word.LoadImages'
        ExpectedCases = @(
            foreach ($imageCount in @(50, 500)) {
                foreach ($accessMode in @('ReadOnly', 'ReadWrite')) { "Load|ImageCount=$imageCount&AccessMode=$accessMode" }
            }
        )
    }
    wordmarkdown = [pscustomobject]@{
        Area = 'word'
        Title = 'Markdown and DOCX conversion'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only conversion scaling measurement; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordMarkdownConversionBenchmarks*'
        ComparisonId = 'word-markdown-conversion'
        Suite = 'OfficeIMO.Word.MarkdownConversion'
        ExpectedCases = @(New-WordExpectedCases -Methods @('MarkdownToWord', 'WordToMarkdown', 'WordToMarkdownChunks') -Parameter 'SectionCount' -Values @(100, 400))
    }
    wordruns = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX run enumeration'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only API scaling measurement; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordRunEnumerationBenchmarks*'
        ComparisonId = 'word-run-enumeration'
        Suite = 'OfficeIMO.Word.RunEnumeration'
        ExpectedCases = @(New-WordExpectedCases -Methods @('EnumerateRuns') -Parameter 'ParagraphCount' -Values @(1000, 4000))
    }
    wordsections = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX document-wide collections across sections'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only API variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordSectionCollectionBenchmarks*'
        ComparisonId = 'word-section-collections'
        Suite = 'OfficeIMO.Word.SectionCollections'
        ExpectedCases = @(New-WordExpectedCases -Methods @('ParagraphsAndTables', 'HyperlinkParagraphs', 'Elements', 'FieldsHyperlinksAndBookmarks', 'ControlsNotesAndNestedTables') -Parameter 'SectionCount' -Values @(25, 100))
    }
    wordworkflows = [pscustomobject]@{
        Area = 'word'
        Title = 'DOCX workflows: fields, mail merge, HTML export, comparison and macros'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only API variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*WordWorkflowBenchmarks*'
        ComparisonId = 'word-workflows'
        Suite = 'OfficeIMO.Word.Workflows'
        ExpectedCases = @(New-WordExpectedCases -Methods @('UpdateFields', 'ExecuteMailMerge', 'ExportHtml', 'LoadDocument', 'ExportLoadedHtml', 'CompareStructure', 'InspectMacroSignatures') -Parameter 'ItemCount' -Values @(100, 1000))
    }
}
