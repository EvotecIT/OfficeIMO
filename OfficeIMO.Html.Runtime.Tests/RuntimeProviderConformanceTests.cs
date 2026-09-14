using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Conformance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeProviderConformanceTests {
    [Fact]
    public async Task ProcessRuntimePassesTheProviderNeutralConformanceSuite() {
        var host = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

        HtmlRuntimeConformanceReport report = await HtmlRuntimeConformanceSuite.RunAsync(host);

        Assert.True(report.Passed, string.Join(Environment.NewLine, report.Cases.Where(result => !result.Passed).Select(result => result.Id + ": " + result.Error)));
        Assert.Equal(3, report.Cases.Count);
    }
}
