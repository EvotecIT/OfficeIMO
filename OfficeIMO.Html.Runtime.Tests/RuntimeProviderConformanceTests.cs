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
        Assert.Equal(7, report.Cases.Count);
        Assert.Equal(45, report.Cases.Sum(item => item.RequiredAssertions));
        Assert.All(HtmlRuntimeQualificationCatalog.All, manifest => {
            HtmlRuntimeProviderExpectation? provider = manifest.Providers.SingleOrDefault(item => item.Id == report.Provider.Id);
            if (provider != null) Assert.True(HtmlRuntimeQualificationCatalog.EvaluateProvider(manifest, report).Passed, manifest.Id);
        });

        HtmlRuntimeQualificationManifest scripted = HtmlRuntimeQualificationCatalog.Get("scripted-document-v1");
        HtmlRuntimeConformanceReport selected = await HtmlRuntimeConformanceSuite.RunAsync(host, scripted);
        Assert.Equal(2, selected.Cases.Count);
        Assert.Equal(11, selected.Cases.Sum(item => item.RequiredAssertions));
        Assert.True(HtmlRuntimeQualificationCatalog.EvaluateProvider(scripted, selected).Passed);

        HtmlRuntimeQualificationManifest automation = HtmlRuntimeQualificationCatalog.Get("programmatic-automation-v1");
        Assert.False(HtmlRuntimeQualificationCatalog.Evaluate(automation, report).Passed);
        Assert.False(HtmlRuntimeQualificationCatalog.Evaluate(automation, report,
            new HtmlRuntimeConsumerQualificationResult {
                ProfileId = automation.Id,
                ConsumerId = automation.Consumers.Single().Id,
                RequiredWorkflows = 1,
                PassedWorkflows = 0,
                FailedWorkflows = 1,
                Passed = true
            }).Passed);
    }
}
