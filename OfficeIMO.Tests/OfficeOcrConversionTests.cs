using OfficeIMO.Excel;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Pdf.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeOcrConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RecognizedTableExportsKeepItsCellsAndSourceGeometry(bool asynchronous) {
        PdfOcrMergeResult source = await RecognizeTableAsync();
        PdfLogicalTable table = Assert.Single(source.Document.Tables);
        var rows = table.Rows.Select(row => row.ToArray()).ToArray();
        var bounds = table.VisualBounds;
        using var wordStream = new MemoryStream();
        using var excelStream = new MemoryStream();
        using var slideStream = new MemoryStream();
        using var htmlStream = new MemoryStream();
        OfficeConversionReport wordReport = asynchronous ? await source.SaveAsWordAsync(wordStream) : source.SaveAsWord(wordStream);
        OfficeConversionReport excelReport = asynchronous ? await source.SaveAsExcelAsync(excelStream) : source.SaveAsExcel(excelStream);
        OfficeConversionReport slideReport = asynchronous ? await source.SaveAsPowerPointAsync(slideStream) : source.SaveAsPowerPoint(slideStream);
        OfficeConversionReport htmlReport = asynchronous ? await source.SaveAsHtmlAsync(htmlStream) : source.SaveAsHtml(htmlStream);
        Assert.True(wordStream.CanWrite && excelStream.CanWrite && slideStream.CanWrite && htmlStream.CanWrite);
        Assert.Equal("Html", htmlReport.Target);
        Assert.Contains("Alpha", Encoding.UTF8.GetString(htmlStream.ToArray()));
        Assert.Equal("Word", wordReport.Target);
        Assert.Equal("Excel", excelReport.Target);
        Assert.Equal("PowerPoint", slideReport.Target);
        using var word = WordDocument.Load(new MemoryStream(wordStream.ToArray()));
        WordTable wordTable = Assert.Single(word.Tables);
        Assert.Contains("Alpha", wordTable.Rows[1].Cells[0].Paragraphs.Select(paragraph => paragraph.Text));
        using var excel = ExcelDocument.Load(new MemoryStream(excelStream.ToArray()));
        Assert.True(excel.Sheets[0].TryGetCellText(2, 1, out string header));
        Assert.Equal("Product", header);
        Assert.True(excel.Sheets[0].TryGetCellText(3, 1, out string cell));
        Assert.Equal("Alpha", cell);
        using var presentation = PowerPointPresentation.Load(new MemoryStream(slideStream.ToArray()));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.Tables), table => table.Rows > 1 && table.GetCell(1, 0).Text == "Alpha");
        Assert.Same(bounds, table.VisualBounds);
        Assert.Equal(rows.Length, table.Rows.Count);
        for (int row = 0; row < rows.Length; row++) Assert.Equal(rows[row], table.Rows[row]);
    }

    [Fact]
    public async Task LossPolicyAndCancellationLeaveDestinationUntouched() {
        PdfOcrMergeResult source = await RecognizeTableAsync();
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, writable: true);
        Assert.Throws<OfficeConversionException>(() => source.SaveAsWord(destination, new WordPdfImportOptions { FailOnLoss = true }));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.SaveAsExcelAsync(destination, cancellationToken: cancellation.Token));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    private static async Task<PdfOcrMergeResult> RecognizeTableAsync() {
        string[][] rows = { new[] { "Product", "Count" }, new[] { "Alpha", "24" }, new[] { "Beta", "12" }, new[] { "Gamma", "36" } };
        var spans = rows.SelectMany((row, index) => row.Select((text, column) => new OcrTextSpan {
            Text = text, Confidence = 1D, Level = OcrTextSpanLevel.Word, LineId = index.ToString(),
            CoordinateUnit = OcrCoordinateUnit.Normalized,
            Region = new OcrRegion { X = column == 0 ? .2D : .5D, Y = .2D + index * .04D, Width = .1D, Height = .02D }
        })).ToArray();
        var engine = new DelegateOcrEngine("table", (_, _) => Task.FromResult(new OcrResult { Spans = spans }));
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAADklEQVR4nGP4CgYMEAoAUHILfRDviCcAAAAASUVORK5CYII=");
        using var generated = PdfDocument.Create(builder => builder.Content(content => content.Image(png, 200, 100)));
        using var source = PdfDocument.Load(new MemoryStream(generated.ToBytes()));
        return await source.ReadWithOcrAsync(engine, new PdfOcrMergeOptions { ReconstructLayout = true, Dpi = 72 });
    }
}
