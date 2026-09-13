using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFrameworkTests {
    [Fact]
    public async Task PreactFetchStateInputStorageAndRemountProduceIndependentConvertibleCaptures() {
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        var origin = new Uri("https://application.example/");
        var resources = new[] { "preact.umd.js", "hooks.umd.js", "report.js" }.Select(name =>
            HtmlRuntimeResource.FromText(new Uri(origin, name), File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Preact", name)), "text/javascript")).ToList();
        resources.Add(HtmlRuntimeResource.FromText(new Uri(origin, "data.json"), "[{\"name\":\"North\",\"value\":24},{\"name\":\"South\",\"value\":18}]", "application/json"));
        await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = origin,
            Html = "<!doctype html><div id='app'></div><script src='/preact.umd.js'></script><script src='/hooks.umd.js'></script><script src='/report.js'></script>",
            Resources = resources,
            Scripts = new[] { "localStorage.setItem('report:name','Monthly');mountReport()" }
        });
        var before = await session.CaptureAsync("document.querySelector('#total').textContent==='Total: 42'");
        Assert.Equal("Report: Monthly", before.Document.QuerySelector("#report-name")!.TextContent);
        var region = session.Locator(HtmlLocatorQuery.ByAccessibleName("Region"));
        await region.SelectOptionsAsync(new[] { "South" });
        await session.Locator("#total").WaitForTextAsync("Total: 18");
        var filtered = await session.CaptureAsync();
        await session.ExecuteAsync("document.querySelector('#region').selectedIndex=1;document.querySelector('#region').dispatchEvent(new Event('change',{bubbles:true}))");
        await session.Locator("#total").WaitForTextAsync("Total: 24");
        await region.SelectOptionsAsync(new[] { "all" });
        await session.Locator("#total").WaitForTextAsync("Total: 42");
        await session.Locator(HtmlLocatorQuery.ByText("Add adjustment")).ClickAsync();
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Report name")).FillAsync("Quarterly");
        await session.WaitForAsync("document.querySelector('#adjustments').textContent==='Adjustments: 1' && localStorage.getItem('report:name')==='Quarterly'");
        var after = await session.CaptureAsync();
        Assert.Equal("Report: Quarterly", after.Document.QuerySelector("#report-name")!.TextContent);
        Assert.Equal("Quarterly", after.Document.QuerySelector("#name")!.FormState!.Value);
        Assert.True((await session.EvaluateAsync("mutationTypes.includes('childList')")).GetBoolean());
        await session.ExecuteAsync("unmountReport();mountReport()");
        var restored = await session.CaptureAsync("document.querySelector('#total').textContent==='Total: 42'");
        Assert.Equal("Report: Quarterly", restored.Document.QuerySelector("#report-name")!.TextContent);
        Assert.Equal("Adjustments: 0", restored.Document.QuerySelector("#adjustments")!.TextContent);
        await session.DisposeAsync();
        Assert.Equal("Total: 18", filtered.Document.QuerySelector("#total")!.TextContent);
        Assert.Contains("Total: 18", PdfReadDocument.Open(HtmlConversionDocument.FromDocument(filtered.Document).ToPdfBytes()).ExtractText());
        Assert.Equal("Report: Monthly", before.Document.QuerySelector("#report-name")!.TextContent);
        Assert.Equal("Monthly", before.Document.QuerySelector("#name")!.FormState!.Value);
        Assert.Equal("Adjustments: 1", after.Document.QuerySelector("#adjustments")!.TextContent);
        var conversion = HtmlConversionDocument.FromDocument(after.Document, new() { BaseUri = after.DocumentUrl });
        Assert.Contains("Report: Quarterly", conversion.ToMarkdown());
        string pdfText = PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText();
        Assert.Contains("Total: 42", pdfText);
        Assert.Contains("Report: Quarterly", pdfText);
    }
}
