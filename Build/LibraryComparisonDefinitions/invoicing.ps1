param([ValidateSet('net8.0', 'net10.0')] [string] $Framework = 'net10.0')

# OfficeIMO's electronic invoice operations. Rules validation (InvoiceRulesValidationBenchmarks) needs a pinned local
# Java, Saxon and rule-bundle authority, so the runner selects only the self-contained operations by category.
[ordered]@{
    invoicing = [pscustomobject]@{
        Area = 'invoicing'
        Title = 'Electronic invoice operations'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only invoice operations; not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.Invoicing.Benchmarks\OfficeIMO.Invoicing.Benchmarks.csproj'
        RunnerArguments = @('--anyCategories', 'xml-write', 'xml-read', 'pdf-generation', '--buildTimeout', '600')
        Filter = '*Invoice*Benchmarks*'
        ComparisonId = 'officeimo-invoice-operations'
        Suite = 'OfficeIMO.Invoicing.Operations'
        ExpectedCases = @('WriteXml', 'ReadXml', 'GeneratePdf')
    }
}
