using OfficeIMO.Epub;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class EpubPdfRunningContentTests {
    [Fact]
    public void StyledConversionReportsUnsupportedStructuredRunningContentAndKeepsCssMarginText() {
        var book = EpubPublication.Create("Running content", "en", "urn:officeimo:epub-pdf-running-content");
        book.AddStylesheet("print", "EPUB/book.css", "@page { margin: 36pt; @bottom-center { content: 'CSS page ' counter(page) } }");
        book.AddChapter("first", "EPUB/first.xhtml", "First chapter", "<h1>First chapter</h1>", ["print"]);
        using var callerPdf = PdfDocument.Create();
        callerPdf.Header(header => header.FirstPageContent(content => content.AddParagraph("Caller header"), OfficeLength.Points(12D)));
        var options = new OfficeConversionOptions { Html = new HtmlPdfExportOptions { PdfOptions = callerPdf.Options } };
        using var source = new MemoryStream(book.Write().Bytes);
        using var output = new MemoryStream();

        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf, options);
        OfficeDiagnostic omission = Assert.Single(report.Diagnostics, item => item.Code == "SectionRunningContentUnavailable");
        Assert.Equal(OfficeConversionLossKind.Omission, omission.LossKind);
        using var pdf = PdfDocument.Load(output.ToArray(), new PdfLoadOptions { IncludeArtifactText = true });
        string text = pdf.ExtractText();
        Assert.Contains("CSS page 1", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Caller header", text, StringComparison.Ordinal);

        options.FailOnLoss = true;
        source.Position = 0;
        using var strictOutput = new MemoryStream();
        OfficeConversionException failure = Assert.Throws<OfficeConversionException>(() =>
            OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, strictOutput, OfficeDocumentFormat.Pdf, options));
        Assert.Contains(failure.Report.Diagnostics, item => item.Code == omission.Code);
        Assert.Equal(0, strictOutput.Length);
    }
}
