using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;

if (args.Length != 2) throw new ArgumentException("Supply the deployed worker DLL path and an output directory.");
var runtime = new HtmlProcessRuntimeProvider(args[0], AngleSharpDomServices.Instance);
var documentUrl = new Uri("https://report.example/");
await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
    Html = await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "scripted-report.html")),
    DocumentUrl = documentUrl,
    Resources = new[] {
        HtmlRuntimeResource.FromText(new Uri(documentUrl, "scripted-report.css"), await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "scripted-report.css")), "text/css"),
        HtmlRuntimeResource.FromText(new Uri(documentUrl, "scripted-report.js"), await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "scripted-report.js")), "text/javascript"),
        HtmlRuntimeResource.FromText(new Uri(documentUrl, "scripted-report.json"), await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "scripted-report.json")), "application/json")
    },
    ReadyExpression = "window.reportReady === true"
});
var initial = await session.CaptureAsync("true");
await session.ExecuteAsync("document.querySelector('#prepare').click();");
await session.WaitForAsync("window.reportReady === true");
var captured = await session.CaptureAsync();
await session.DisposeAsync();
if (initial.Document.OuterHtml == captured.Document.OuterHtml) throw new InvalidOperationException("The report did not change between captures.");
var document = HtmlConversionDocument.FromDocument(captured.Document, new() { BaseUri = captured.DocumentUrl });
var resources = captured.Resources.ToDictionary(resource => resource.Url.AbsoluteUri, StringComparer.Ordinal);
HtmlRenderResourceResolver resolver = (request, _) => Task.FromResult(resources.TryGetValue(request.Uri.AbsoluteUri, out var resource)
    ? new HtmlResolvedResource(resource.Content, resource.ContentType, resource.FinalUrl, resource.RedirectCount) : null);
var renderOptions = new HtmlRenderOptions { ResourceResolver = resolver };
string output = Path.GetFullPath(args[1]);
Directory.CreateDirectory(output);
await File.WriteAllTextAsync(Path.Combine(output, "report.html"), captured.Document.OuterHtml);
foreach (string asset in new[] { "scripted-report.css", "scripted-report.js", "scripted-report.json" })
    await File.WriteAllBytesAsync(Path.Combine(output, asset), resources[new Uri(documentUrl, asset).AbsoluteUri].Content);
await File.WriteAllTextAsync(Path.Combine(output, "report.md"), document.ToMarkdown());
await File.WriteAllTextAsync(Path.Combine(output, "report.svg"), await document.ToSvgAsync(renderOptions));
await File.WriteAllBytesAsync(Path.Combine(output, "report.png"), await document.ToPngAsync(renderOptions));
byte[] pdf = await document.ToPdfBytesAsync(new HtmlToPdfOptions { ResourceResolver = resolver });
await File.WriteAllBytesAsync(Path.Combine(output, "report.pdf"), pdf);
if (!PdfReadDocument.Open(pdf).ExtractText().Contains("Total: 42")) throw new InvalidOperationException("The script-generated report is missing from the PDF.");
Console.WriteLine($"Captured with {captured.ProviderId}. Wrote HTML, Markdown, SVG, PNG and searchable PDF to {output}.");
