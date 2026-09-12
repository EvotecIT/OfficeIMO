using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;

if (args.Length != 2) throw new ArgumentException("Supply the deployed worker DLL path and an output directory.");
var runtime = new HtmlProcessRuntimeProvider(args[0], AngleSharpDomServices.Instance);
await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
    Html = await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "scripted-report.html")),
    ReadyExpression = "window.reportReady === true"
});
var initial = await session.CaptureAsync("true");
await session.ExecuteAsync("document.querySelector('#prepare').click();");
await session.WaitForAsync("window.reportReady === true");
var captured = await session.CaptureAsync();
await session.DisposeAsync();
if (initial.Document.OuterHtml == captured.Document.OuterHtml) throw new InvalidOperationException("The report did not change between captures.");
var document = HtmlConversionDocument.FromDocument(captured.Document);
string output = Path.GetFullPath(args[1]);
Directory.CreateDirectory(output);
await File.WriteAllTextAsync(Path.Combine(output, "report.html"), captured.Document.OuterHtml);
await File.WriteAllTextAsync(Path.Combine(output, "report.md"), document.ToMarkdown());
await File.WriteAllTextAsync(Path.Combine(output, "report.svg"), document.ToSvg());
await File.WriteAllBytesAsync(Path.Combine(output, "report.png"), document.ToPng());
byte[] pdf = document.ToPdfBytes();
await File.WriteAllBytesAsync(Path.Combine(output, "report.pdf"), pdf);
if (!PdfReadDocument.Open(pdf).ExtractText().Contains("Total: 42")) throw new InvalidOperationException("The script-generated report is missing from the PDF.");
Console.WriteLine($"Captured with {captured.ProviderId}. Wrote HTML, Markdown, SVG, PNG and searchable PDF to {output}.");
