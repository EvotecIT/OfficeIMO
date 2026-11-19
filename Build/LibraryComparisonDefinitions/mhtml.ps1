param([ValidateSet('net8.0', 'net10.0')] [string] $Framework = 'net10.0')

# MHTML archive reading and writing against MimeKit. Each lane's output is read back through both OfficeIMO and
# MimeKit and compared with the corpus: subject, root Content-ID and location, root HTML, and decoded resource bytes.
[ordered]@{
    mhtmlread = [pscustomobject]@{
        Area = 'mhtml'
        Title = 'MHTML archive reading'
        Libraries = @('OfficeIMO', 'MimeKit')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'The MimeKit lane also parses the root HTML with AngleSharp, and the catalog cannot name a combined MimeKit and AngleSharp lane yet.'
        CatalogEligible = $false
        Project = 'OfficeIMO.Mhtml.Benchmarks.Comparisons\OfficeIMO.Mhtml.Benchmarks.Comparisons.csproj'
        Filter = '*MhtmlReadBenchmarks*'
        ComparisonId = 'mhtml-read'
        Suite = 'OfficeIMO.Mhtml.Read'
        ExpectedCases = @(
            foreach ($scale in @('Small', 'Normal', 'Large')) {
                foreach ($engine in @('OfficeIMO', 'MimeKit')) {
                    "$engine|Scale=$scale"
                }
            }
        )
    }
    mhtmlwrite = [pscustomobject]@{
        Area = 'mhtml'
        Title = 'MHTML archive writing'
        Libraries = @('OfficeIMO', 'MimeKit')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.Mhtml.Benchmarks.Comparisons\OfficeIMO.Mhtml.Benchmarks.Comparisons.csproj'
        Filter = '*MhtmlWriteBenchmarks*'
        ComparisonId = 'mhtml-write'
        Suite = 'OfficeIMO.Mhtml.Write'
        ExpectedCases = @(
            foreach ($scale in @('Small', 'Normal', 'Large')) {
                foreach ($engine in @('OfficeIMO', 'MimeKit')) {
                    "$engine|Scale=$scale"
                }
            }
        )
    }
}
