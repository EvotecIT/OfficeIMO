using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeStandaloneApplicationTests {
    private static readonly Uri Start = new("https://application.example/index.html");

    [Fact]
    public async Task WebApplicationProfileCompletesAStandaloneTwoPageWorkflow() {
        string Read(string name) => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "StandaloneApplication", name));
        HtmlRuntimeResource Text(string relative, string contentType) =>
            HtmlRuntimeResource.FromText(new Uri(Start, relative), Read(relative), contentType);
        var reviewUrl = new Uri(Start, "/review?title=Quarterly&region=South");
        var resources = new List<HtmlRuntimeResource> {
            Text("app.js", "text/javascript"),
            Text("view.js", "text/javascript"),
            Text("app.css", "text/css"),
            Text("theme.css", "text/css"),
            Text("data.json", "application/json"),
            new(new Uri(Start, "health.txt"), Encoding.UTF8.GetBytes(Read("health.txt")), "text/plain", statusCode: 503, statusText: "Service Unavailable"),
            HtmlRuntimeResource.FromText(reviewUrl, Read("review.html"), "text/html; charset=utf-8"),
            Text("review.js", "text/javascript")
        };
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        await using var session = await runtime.OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = Start,
            Html = Read("index.html"),
            Resources = resources,
            ViewportWidth = 360,
            ViewportHeight = 180,
            ReadyExpression = "window.applicationReady===true"
        });

        await session.WaitForAsync("window.applicationReady===true");
        var heading = session.Locator("h1");
        await heading.WaitForTextAsync("Application dashboard");
        await session.Locator("#total").WaitForTextAsync("Total: 49");
        Assert.Equal("Service degraded (503)", (await session.Locator("#health").InspectAsync()).Text);
        Assert.True((await session.EvaluateAsync("applicationMutations>0 && document.readyState==='complete'")).GetBoolean());

        var title = session.Locator(HtmlLocatorQuery.ByAccessibleName("Report title"));
        await title.FillAsync("quarterly");
        await title.SetSelectionAsync(0, 1);
        await title.PressAsync("Q");
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Region")).SelectOptionsAsync(new[] { "South" });
        await session.Locator("#total").WaitForTextAsync("Total: 25");
        Assert.Equal("Quarterly", (await title.InspectAsync()).Value);
        Assert.Equal(Start.AbsoluteUri, (await session.EvaluateAsync("document.URL")).GetString());
        Assert.Equal(Start.AbsoluteUri, (await session.EvaluateAsync("document.baseURI")).GetString());
        Assert.Equal(Start.AbsoluteUri, (await session.EvaluateAsync("location.href")).GetString());

        var submit = session.Locator(HtmlLocatorQuery.ByAccessibleName("Prepare review"));
        var submitState = await submit.InspectAsync();
        Assert.True(submitState.IsVisible);
        Assert.False(submitState.IsInViewport);
        Assert.Equal(180D, submitState.BoundingBox!.Width);
        await submit.ClickAsync();

        await heading.WaitForTextAsync("Review report");
        await session.Locator("#review-total").WaitForTextAsync("Total: 25");
        Assert.Equal(reviewUrl, (await session.CaptureAsync("window.reviewReady===true")).DocumentUrl);
        Assert.Equal("Quarterly / South", (await session.Locator("#selection").InspectAsync()).Text);
        Assert.Equal("Stored: Quarterly / South", (await session.Locator("#stored").InspectAsync()).Text);
        Assert.Equal("Lifecycle: pagehide=false, unload=yes", (await session.Locator("#lifecycle").InspectAsync()).Text);
        var review = await session.CaptureAsync("window.reviewReady===true");

        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Back to dashboard")).ClickAsync();
        await heading.WaitForTextAsync("Application dashboard");
        await session.WaitForAsync("window.applicationReady===true");
        Assert.Equal("Quarterly", (await title.InspectAsync()).Value);
        Assert.Equal(new[] { "South" }, (await session.Locator("#region").InspectAsync()).SelectedValues);
        Assert.Equal("Total: 25", (await session.Locator("#total").InspectAsync()).Text);

        await session.DisposeAsync();
        Assert.Equal("Review report", review.Document.QuerySelector("h1")!.TextContent);
        var conversion = HtmlConversionDocument.FromDocument(review.CreateStandaloneDocument());
        Assert.Contains("Quarterly / South", conversion.ToMarkdown());
        string pdfText = PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText();
        Assert.Contains("Review report", pdfText);
        Assert.Contains("Total: 25", pdfText);
    }
}
