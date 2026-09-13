using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeObservedReportTests {
    [Fact]
    public async Task ModuleObserverReportSurvivesActionsSessionDisposalAndConversion() {
        var page = new Uri("https://reports.example/app/index.html");
        string Read(string name) => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "ObservedReport", name));
        var runtime = new HtmlProcessRuntimeProvider(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        await using var session = await runtime.OpenTrustedAsync(new() {
            DocumentUrl = page,
            Html = Read("index.html"),
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(page, "report.js"), Read("report.js"), "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(page, "items.json"), Read("items.json"), "application/json")
            }
        });
        await session.Locator("#total").WaitForTextAsync("Total: 42");
        Assert.Equal("complete", (await session.EvaluateAsync("reportReadyState")).GetString());
        Assert.Equal("before,observer,after", (await session.EvaluateAsync("reportOrder.join(',')")).GetString());
        var before = await session.CaptureAsync();
        await session.Locator("#update").ClickAsync();
        await session.Locator("#total").WaitForTextAsync("Total: 44");
        var after = await session.CaptureAsync();
        await session.DisposeAsync();

        Assert.Equal("Total: 42", before.Document.QuerySelector("#total")!.TextContent);
        Assert.Equal("Total: 44", after.Document.QuerySelector("#total")!.TextContent);
        Assert.Equal("20", after.Document.QuerySelectorAll("[data-amount]")[1].TextContent);
        Assert.Contains(after.Resources, resource => resource.Url.AbsolutePath == "/app/items.json");
        var conversion = HtmlConversionDocument.FromDocument(after.CreateStandaloneDocument());
        Assert.Contains("Total: 44", conversion.ToMarkdown());
        var text = PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText();
        Assert.Contains("Total: 44", text);
        Assert.Contains("Services", text);
        Assert.Contains("Materials", text);
    }
}
