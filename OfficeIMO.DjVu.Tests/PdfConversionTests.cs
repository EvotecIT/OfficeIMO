using System.Threading.Tasks;
using OfficeIMO.DjVu.Pdf;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.DjVu.Tests;

public sealed class PdfConversionTests {
    [Fact]
    public void StoredUnicodeIsSearchableWithPhysicalRotatedPageSize() {
        var document = DjVuDocument.Load(ReaderAdapterTests.Fixture("unicode.djvu"));
        var result = document.ToPdfDocumentResult();
        var report = Assert.IsType<DjVuPdfConversionReport>(Assert.Single(result.SourceConversionReports));
        var pageReport = Assert.Single(report.Pages);
        Assert.Equal(DjVuPdfTextSource.StoredText, pageReport.TextSource);
        Assert.Null(pageReport.Recognition);
        Assert.Equal(document.Pages[0].GetText().Text.Length, pageReport.TextCharacters);
        Assert.Equal(96 * 72.0 / 300, pageReport.WidthPoints, 8);
        Assert.Equal(128 * 72.0 / 300, pageReport.HeightPoints, 8);
        var pdf = PdfReadDocument.Open(result.ToBytes());
        Assert.Contains("Zażółć 😀", pdf.ExtractText());
        Assert.Single(pdf.Pages);
        using var output = new MemoryStream(new byte[100_000], true);
        document.SaveAsPdf(output);
        Assert.True(output.CanWrite);
        Assert.Equal(0, output.Position);
        Assert.True(output.Length < 100_000);
        Assert.Contains("Zażółć", PdfReadDocument.Open(output.ToArray()).ExtractText());
    }

    [Fact]
    public async Task ExplicitOcrRunsOnlyForSelectedAbsentAndEmptyTextAndPreservesProvenance() {
        var document = DjVuDocument.Load(ReaderAdapterTests.Fixture("reader-book.djvu"));
        var calls = new System.Collections.Generic.List<int>();
        var engine = new DelegateOcrEngine("fixture-engine", (request, token) => {
            calls.Add(request.PageNumber!.Value);
            Assert.Equal("image/png", request.MediaType);
            Assert.True(request.Payload.Length > 8);
            return Task.FromResult(new OcrResult { Text = "Recognized " + request.PageNumber,
                Provider = "fixture-provider", Model = "fixture-v1", Confidence = .99,
                Spans = new[] { new OcrTextSpan { Sequence = 0, Text = "Recognized " + request.PageNumber, Level = OcrTextSpanLevel.Word,
                    CoordinateUnit = OcrCoordinateUnit.Normalized, Region = new OcrRegion { X = .1, Y = .2, Width = .6, Height = .15 }, Confidence = .99 } } });
        }, new OcrEngineCapabilities { SupportsWordSpans = true });
        var result = await document.ToPdfDocumentResultAsync(new DjVuToPdfOptions { OcrEngine = engine, PageNumbers = new[] { 4, 2, 3, 1 } });
        Assert.Equal(new[] { 3, 1 }, calls);
        var report = Assert.IsType<DjVuPdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal(new[] { 4, 2, 3, 1 }, report.Pages.Select(p => p.SourcePageNumber));
        Assert.Equal(new[] { DjVuPdfTextSource.None, DjVuPdfTextSource.StoredText, DjVuPdfTextSource.Ocr, DjVuPdfTextSource.Ocr }, report.Pages.Select(p => p.TextSource));
        Assert.Null(report.Pages[1].Recognition);
        Assert.Equal("fixture-provider", report.Pages[2].Recognition!.Provider);
        Assert.Equal(.99, report.Pages[2].Recognition!.Confidence!.Value, 8);
        Assert.NotNull(report.OcrReport);
        Assert.Contains(report.FidelityDiagnostics, d => d.Code == "djvu.pdf.stored-text-corrupt");
        var pdf = PdfReadDocument.Open(result.ToBytes());
        Assert.Equal(4, pdf.Pages.Count);
        Assert.Contains("Zażółć 😀", pdf.ExtractText());
        Assert.Contains("Recognized 3", pdf.ExtractText());
        Assert.Contains("Recognized 1", pdf.ExtractText());
        calls.Clear();
        await document.ToPdfDocumentResultAsync(new DjVuToPdfOptions { OcrEngine = engine, PageNumbers = new[] { 2, 4 } });
        Assert.Empty(calls);
        Assert.Throws<InvalidOperationException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { OcrEngine = engine }));
    }

    [Fact]
    public async Task GeometryFreeOcrAndInvalidOutputCannotProduceInventedSearchLocations() {
        var document = DjVuDocument.Load(ReaderAdapterTests.Fixture("palette.djvu"));
        int calls = 0;
        var engine = new DelegateOcrEngine("fixture", (request, token) => { calls++; return Task.FromResult(new OcrResult { Text = "No geometry" }); });
        var settings = new DjVuToPdfOptions { OcrEngine = engine };
        using var readOnly = new MemoryStream(Array.Empty<byte>(), false);
        await Assert.ThrowsAsync<ArgumentException>(() => document.SaveAsPdfAsync(readOnly, settings));
        await Assert.ThrowsAsync<ArgumentException>(() => document.SaveAsPdfAsync(" ", settings));
        Assert.Equal(0, calls);
        var result = await document.ToPdfDocumentResultAsync(settings);
        var report = Assert.IsType<DjVuPdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal(DjVuPdfTextSource.None, report.Pages[0].TextSource);
        Assert.Contains(report.FidelityDiagnostics, d => d.Code == "djvu.pdf.ocr-geometry-missing");
        Assert.Equal(string.Empty, PdfReadDocument.Open(result.ToBytes()).ExtractText().Trim());
    }

    [Fact]
    public void PageTextImageAndSerializationLimitsAndCancellationStayExplicit() {
        var document = DjVuDocument.Load(ReaderAdapterTests.Fixture("reader-book.djvu"));
        Assert.Throws<DjVuResourceLimitException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { MaxPages = 1 }));
        Assert.Throws<DjVuResourceLimitException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { PageNumbers = new[] { 2 }, MaxTextCharacters = 1 }));
        Assert.Throws<ArgumentException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { PageNumbers = new[] { 2, 2 } }));
        Assert.Throws<ArgumentException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { PageNumbers = Array.Empty<int>() }));
        Assert.Throws<OperationCanceledException>(() => document.ToPdfBytes(cancellationToken: new CancellationToken(true)));
        Assert.Throws<InvalidDataException>(() => document.ToPdfBytes(new DjVuToPdfOptions { PageNumbers = new[] { 2 }, MaxPdfBytes = 128 }));
        Assert.Throws<OfficeIMO.Drawing.OfficeImageExportBatchLimitException>(() => document.ToPdfDocumentResult(new DjVuToPdfOptions { MaxTotalImageBytes = 32 }));
    }
}
