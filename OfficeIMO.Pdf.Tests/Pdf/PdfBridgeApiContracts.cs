using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Reader;
using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfBridgeApiContracts {
    [Fact]
    public void OpenedDocumentAdaptersPropagateTheCanonicalSemanticPageLimit() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(pdf));
        var readOptions = new PdfLoadOptions {
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
        };

        Assert.Throws<PdfReadLimitException>(() =>
            document.ToWord(new WordPdfImportOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToExcel(new ExcelPdfImportOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToRtfDocument(new PdfToRtfOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToHtml(new PdfToHtmlOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToPowerPoint(new PowerPointPdfImportOptions { ReadOptions = readOptions }));
        using var stream = new MemoryStream(pdf, writable: false);
        ReaderAdmissionLimitException error = Assert.Throws<ReaderAdmissionLimitException>(() =>
            PdfReaderAdapter.ReadDocument(stream, pdfOptions: new ReaderPdfOptions { ReadOptions = readOptions }));
        Assert.Equal(PdfReadLimitKind.Pages, Assert.IsType<PdfReadLimitException>(error.InnerException).Kind);
        Assert.Equal(1, error.Maximum); Assert.Equal(2, error.Actual);
    }

    [Fact]
    public void WordImportUsesCanonicalPageSelectionWithinTheConfiguredLimit() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page marker"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page marker"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(pdf));
        var options = new WordPdfImportOptions {
            ReadOptions = new PdfLoadOptions {
                PageSelection = PdfPageSelection.From(2),
                Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
            }
        };

        using OfficeIMO.Word.WordDocument word = document.ToWord(options);
        string text = string.Join(" ", word.Paragraphs.Select(static paragraph => paragraph.Text));

        Assert.Contains("Second page marker", text, StringComparison.Ordinal);
        Assert.DoesNotContain("First page marker", text, StringComparison.Ordinal);
        PdfLoadOptions cloned = options.Clone().ReadOptions!;
        Assert.NotSame(options.ReadOptions, cloned);
        Assert.Equal(options.ReadOptions.PageSelection, cloned.PageSelection);
        Assert.NotNull(cloned.Pipeline);
        Assert.NotNull(options.ReadOptions.Pipeline);
        Assert.Equal(1, cloned.Pipeline!.MaxPages);
        cloned.Pipeline.MaxPages = 2;
        Assert.Equal(1, options.ReadOptions.Pipeline!.MaxPages);
    }

}
