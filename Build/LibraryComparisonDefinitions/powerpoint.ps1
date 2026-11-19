# PowerPoint (PPTX) suites beyond the create and edit comparisons in LibraryComparisonDefinitions.ps1.
param(
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0'
)

$project = 'OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler.csproj'

function New-PowerPointExpectedCases {
    foreach ($slideCount in @(20, 200)) {
        foreach ($engine in @('OfficeIMO', 'ShapeCrawler', 'OpenXmlSdk', 'OneDocument')) { "$engine|SlideCount=$slideCount" }
    }
}

[ordered]@{
    pptxread = [pscustomobject]@{
        Area = 'powerpoint'
        Title = 'PPTX text and table reading'
        Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK', 'OneDocument.Presentations')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        RunnerArguments = @('bench')
        Frameworks = @('net10.0')
        Filter = '*PowerPointReadComparisonBenchmarks*'
        ComparisonId = 'pptx-read'
        Suite = 'OfficeIMO.PowerPoint.Read'
        ExpectedCases = @(New-PowerPointExpectedCases)
    }
    pptxcopyslides = [pscustomobject]@{
        Area = 'powerpoint'
        Title = 'PPTX slide copy and save'
        Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK', 'OneDocument.Presentations')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        RunnerArguments = @('bench')
        Frameworks = @('net10.0')
        Filter = '*PowerPointCopySlidesComparisonBenchmarks*'
        ComparisonId = 'pptx-copy-slides'
        Suite = 'OfficeIMO.PowerPoint.CopySlides'
        ExpectedCases = @(New-PowerPointExpectedCases)
    }
}
