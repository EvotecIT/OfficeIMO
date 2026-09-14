using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime.Conformance;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlRuntimeConformanceRunner {
    internal static async Task<int> RunAsync() {
        await using var host = new ChromiumConformanceRuntimeHost(AngleSharpHtmlParser.Instance);
        HtmlRuntimeConformanceReport report = await HtmlRuntimeConformanceSuite.RunAsync(host);
        Console.WriteLine($"Provider: {report.Provider.Id} {report.Provider.Version}");
        foreach (HtmlRuntimeConformanceCaseResult item in report.Cases)
            Console.WriteLine($"{(item.Passed ? "PASS" : "FAIL")} {item.Id} ({item.Elapsed.TotalMilliseconds:0} ms){(item.Error == null ? string.Empty : ": " + item.Error)}");
        return report.Passed ? 0 : 1;
    }
}
