using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime.Conformance;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlRuntimeConformanceRunner {
    internal static async Task<int> RunAsync() {
        await using var host = new ChromiumConformanceRuntimeHost(AngleSharpHtmlParser.Instance);
        HtmlRuntimeConformanceReport report = await HtmlRuntimeConformanceSuite.RunAsync(host);
        Console.WriteLine($"Provider: {report.Provider.Id} {report.Provider.Version}");
        foreach (HtmlRuntimeConformanceCaseResult item in report.Cases)
            Console.WriteLine($"{(item.Passed ? "PASS" : "FAIL")} {item.Id} ({item.PassedAssertions}/{item.RequiredAssertions} assertions, {item.Elapsed.TotalMilliseconds:0} ms){(item.Error == null ? string.Empty : ": " + item.Error)}");
        HtmlRuntimeQualificationResult[] qualifications = HtmlRuntimeQualificationCatalog.All
            .Where(manifest => manifest.Providers.Any(provider => provider.Id == report.Provider.Id))
            .Select(manifest => HtmlRuntimeQualificationCatalog.EvaluateProvider(manifest, report)).ToArray();
        foreach (HtmlRuntimeQualificationResult result in qualifications)
            Console.WriteLine($"{(result.Passed ? "PASS" : "FAIL")} {result.ProfileId} provider ({result.PassedCases}/{result.RequiredCases} cases, {result.PassedAssertions}/{result.RequiredAssertions} assertions)");
        return report.Passed && qualifications.Length > 0 && qualifications.All(item => item.Passed) ? 0 : 1;
    }
}
