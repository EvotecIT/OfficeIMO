using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfCanvasLogicalTextPrecisionTests {
    [Theory]
    [InlineData(.2214321D, false)]
    [InlineData(.0044321D, false)]
    [InlineData(.2214321D, true)]
    [InlineData(.0044321D, true)]
    public void BoundedActualText_FractionalHeightPreservesAuthoredAdvance(double height, bool embedded) {
        var options = new PdfOptions { CompressContentStreams = false };
        if (embedded) {
            string path = Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf");
            Assert.True(File.Exists(path));
            options.UseFontFamily(new PdfEmbeddedFontFamily("Fractional carrier", File.ReadAllBytes(path)));
        }
        byte[] bytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.ActualText("Logical marker", 10D, 20D, 50D, height,
                paint => paint.Text("PaintOnly", 10D, 20D, 50D, 12D)))
            .ToBytes();
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        PdfTextSpan span = Assert.Single(page.GetTextSpans());
        Assert.Equal("Logical marker", span.Text);
        Assert.Equal(50D, span.Advance, 2);
        Assert.Equal(height, span.FontSize, 6);
        Assert.Equal(10D, span.X, 3);
    }

    [Theory]
    [InlineData(.2214321D)]
    [InlineData(.0044321D)]
    public async Task SearchableOcr_FractionalHeightPreservesIndependentWordAdvance(double height) {
        var engine = new DelegateOcrEngine("fractional-word", (_, _) => Task.FromResult(new OcrResult {
            Provider = "fixture", Language = "eng", Spans = new[] { new OcrTextSpan {
                Text = "Marker", Confidence = .95D, Level = OcrTextSpanLevel.Word,
                CoordinateUnit = OcrCoordinateUnit.Points,
                Region = new OcrRegion { X = 10D, Y = 20D, Width = 50D, Height = height }
            } }
        }));
        byte[] source = PdfDocument.Create(document => document.Page(page => page.Size(300D, 300D))).ToBytes();
        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(engine);
        Assert.Equal(1, result.AddedWordCount);
        Assert.Equal(height, Assert.Single(result.WrittenWords[1]).Height, 6);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(result.Document.ToBytes());
        var letters = independent.GetPage(1).Letters;
        Assert.Equal(6, letters.Count);
        Assert.Equal(10D, letters[0].StartBaseLine.X, 3);
        Assert.Equal(60D, letters[5].EndBaseLine.X, 2);
        Assert.Equal("Marker", PdfReadDocument.Open(result.Document.ToBytes()).ExtractText().Trim());
    }
}
