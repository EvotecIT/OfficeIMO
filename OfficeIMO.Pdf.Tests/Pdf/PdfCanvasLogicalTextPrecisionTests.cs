using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfCanvasLogicalTextPrecisionTests {
    [Theory]
    [InlineData(2812L, .22141732283464566D, false, false)]
    [InlineData(56L, .004409448818897638D, false, false)]
    [InlineData(2812L, .22141732283464566D, true, false)]
    [InlineData(56L, .004409448818897638D, true, false)]
    [InlineData(2812L, .22141732283464566D, true, true)]
    [InlineData(56L, .004409448818897638D, true, true)]
    public void BoundedActualText_FractionalHeightPreservesAuthoredAdvance(long heightEmus, double expectedHeight, bool embedded, bool customProvider) {
        var options = new PdfCreateOptions { CompressContentStreams = false };
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        if (customProvider) options.SetTextShapingProvider(provider);
        if (embedded) {
            string path = Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf");
            Assert.True(File.Exists(path));
            options.UseFontFamily(new PdfEmbeddedFontFamily("Fractional carrier", File.ReadAllBytes(path)));
        }
        byte[] bytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.ActualText("Logical marker", global::OfficeIMO.OfficeLength.Points(10D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(50D), global::OfficeIMO.OfficeLength.Emus(heightEmus),
                paint => paint.Text("PaintOnly", global::OfficeIMO.OfficeLength.Points(10D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(50D), global::OfficeIMO.OfficeLength.Points(12D))))
            .ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        PdfTextSpan span = Assert.Single(page.GetTextSpans());
        Assert.Equal("Logical marker", span.Text);
        Assert.Equal(50D, span.Advance, 2);
        Assert.Equal(expectedHeight, span.FontSize, 6);
        Assert.Equal(10D, span.X, 3);
        if (customProvider) Assert.DoesNotContain(provider.Requests, request => string.IsNullOrWhiteSpace(request.Text));
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
        byte[] source = PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300D), global::OfficeIMO.OfficeLength.Points(300D)))).ToBytes();
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
