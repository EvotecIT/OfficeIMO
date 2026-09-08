using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOcrReconstructionTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task ReconstructsColumnsJoinedByProviderWithoutChangingWordGeometry(int turns) {
        var spans = new List<OcrTextSpan>();
        for (int row = 0; row < 3; row++) {
            foreach (int column in new[] { 0, 1 }) {
                double x = 0.08D + column * 0.5D, y = 0.1D + row * 0.1D;
                string text = (column == 0 ? "Left" : "Right") + row;
                spans.Add(Word(text, x, y, 0.1D, 0.018D, row));
                spans.Add(Word("paragraph", x + 0.11D, y, 0.14D, 0.018D, row));
                spans.Add(Word("ends.", x + 0.26D, y, 0.09D, 0.018D, row));
            }
        }
        var provider = new DelegateOcrEngine("columns", (_, _) => Task.FromResult(new OcrResult { Spans = spans }));
        byte[] source = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 200, 100).ToBytes();
        var options = Options(turns);
        PdfOcrMergeResult preserved = await PdfDocument.Load(source).ReadWithOcrAsync(provider, options);
        options.ReconstructLayout = true;
        Assert.True(options.Clone().ReconstructLayout);
        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider, options);
        Assert.Empty(result.Ocr.Document.Tables);
        const string expected = "Left0 paragraph ends. Left1 paragraph ends. Left2 paragraph ends. Right0 paragraph ends. Right1 paragraph ends. Right2 paragraph ends.";
        Assert.Equal(expected, Normalize(result.Ocr.Text));
        Assert.NotEqual(expected, Normalize(preserved.Text));
        Assert.Equal(preserved.Pages[0].Words.Select(static word => (word.ProviderSequence, word.X, word.Y, word.Width, word.Height)),
            result.Ocr.Pages[0].Words.Select(static word => (word.ProviderSequence, word.X, word.Y, word.Width, word.Height)));
        byte[] searchable = result.Document.ToBytes();
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(source), PdfPageImageRenderer.RenderPageAsPng(searchable));
        Assert.Equal(expected, Normalize(PdfDocument.Load(searchable).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }).Text));
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    public async Task ReconstructsTableRowsInCorrectedFrame(int turns, bool rightToLeft) {
        string[][] expected = rightToLeft
            ? new[] { new[] { "פריט", "כמות" }, new[] { "אלפא", "24" }, new[] { "בטא", "12" }, new[] { "גמא", "36" } }
            : new[] { new[] { "Product", "Count" }, new[] { "Alpha", "24" }, new[] { "Beta", "12" }, new[] { "Gamma", "36" } };
        var spans = expected.SelectMany((row, index) => new[] {
            Word(row[0], rightToLeft ? 0.5D : 0.2D, 0.2D + index * 0.04D, 0.12D, 0.02D, index),
            Word(row[1], rightToLeft ? 0.2D : 0.5D, 0.2D + index * 0.04D, 0.1D, 0.02D, index)
        }).ToArray();
        var provider = new DelegateOcrEngine("table", (_, _) => Task.FromResult(new OcrResult { Spans = spans }));
        byte[] source = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 200, 100).ToBytes();
        var options = Options(turns);
        options.ReconstructLayout = true;
        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider, options);
        PdfLogicalTable table = Assert.Single(result.Ocr.Document.Tables);
        Assert.Equal(expected.Length, table.Rows.Count);
        for (int index = 0; index < expected.Length; index++) Assert.Equal(expected[index], table.Rows[index]);
        Assert.NotNull(table.VisualBounds);
        Assert.True(table.VisualBounds!.Width > 0D);
        Assert.True(table.VisualBounds.Height > 0D);
        PdfLogicalTable readback = Assert.Single(PdfDocument.Load(result.Document.ToBytes())
            .Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }).Tables);
        Assert.Equal(PdfLogicalContentSourceKind.Native, readback.SourceKind);
        for (int index = 0; index < expected.Length; index++) Assert.Equal(expected[index], readback.Rows[index]);
    }

    private static PdfOcrMergeOptions Options(int turns) => new PdfOcrMergeOptions {
        Dpi = 72,
        ScanProcessing = new OfficeScanProcessingOptions { ClockwiseQuarterTurns = turns,
            Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor }
    };

    private static OcrTextSpan Word(string text, double x, double y, double width, double height, int row) => new OcrTextSpan {
        Text = text, Level = OcrTextSpanLevel.Word, Confidence = 1D, LineId = row.ToString(),
        CoordinateUnit = OcrCoordinateUnit.Normalized, Region = new OcrRegion { X = x, Y = y, Width = width, Height = height }
    };

    private static string Normalize(string value) => System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ").Trim();
}
