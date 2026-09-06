using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSearchableOcrReviewTests {
    [Fact]
    public async Task ReviewWritesOnlySelectedWordsAndReportsTheActualLayer() {
        byte[] bytes = Source();
        var source = PdfDocument.Load(bytes);
        var options = new PdfOcrMergeOptions { Dpi = 72 };
        var review = await source.PrepareSearchableOcrAsync(Engine(), options);
        options.Dpi = 300;
        var page = Assert.Single(review.Ocr.Pages);
        Assert.Equal(2, page.Words.Count);
        Assert.Equal(3, page.WordEvidence.Count);
        Assert.Equal(300, review.RenderPage(1).Width);
        Assert.Throws<ArgumentOutOfRangeException>(() => review.RenderPage(2));
        Assert.Equal(bytes, source.ToBytes());
        var result = review.Apply(new[] { page.Words[1] });
        Assert.Equal(1, result.AddedWordCount);
        Assert.Equal(2, result.Ocr.AcceptedWordCount);
        Assert.Equal("Second", Assert.Single(result.WrittenWords[1]).Text);
        Assert.Equal(new[] { 1 }, result.ModifiedPages);
        Assert.Equal("Second", PdfReadDocument.Open(result.Document.ToBytes()).ExtractText().Trim());
        using var independent = UglyToad.PdfPig.PdfDocument.Open(result.Document.ToBytes());
        Assert.Equal("Second", Assert.Single(ActualText(independent.GetPage(1).GetMarkedContents())));
        Assert.Equal(bytes, source.ToBytes());
    }

    [Fact]
    public async Task ReviewRejectsForeignDuplicateAndPolicyRejectedWords() {
        byte[] original = Source();
        var source = PdfDocument.Load(original);
        var review = await source.PrepareSearchableOcrAsync(Engine());
        var other = await source.PrepareSearchableOcrAsync(Engine());
        var word = review.Ocr.Pages[0].Words[0];
        Assert.Throws<ArgumentException>(() => review.Apply(new[] { other.Ocr.Pages[0].Words[0] }));
        Assert.Throws<ArgumentException>(() => review.Apply(new[] { word, word }));
        Assert.Throws<ArgumentException>(() => review.Apply(new[] { review.Ocr.Pages[0].WordEvidence[2].Word }));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => review.ApplyAll(cancellation.Token));
        var unchanged = review.Apply(Array.Empty<PdfRecognizedWord>());
        Assert.False(unchanged.WasModified);
        Assert.Equal(0, unchanged.AddedWordCount);
        Assert.Empty(unchanged.WrittenWords);
        Assert.Equal(original, unchanged.Document.ToBytes());
    }

    private static byte[] Source() => PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();

    private static System.Collections.Generic.IEnumerable<string> ActualText(
        System.Collections.Generic.IEnumerable<UglyToad.PdfPig.Content.MarkedContentElement> elements) {
        foreach (var element in elements) {
            if (!string.IsNullOrEmpty(element.ActualText)) yield return element.ActualText;
            else foreach (string text in ActualText(element.Children)) yield return text;
        }
    }

    private static IOcrEngine Engine() => new DelegateOcrEngine("review-fixture", (_, _) => Task.FromResult(new OcrResult {
        Provider = "fixture", Language = "eng", Spans = new[] {
            Word("First", 20, 0.95), Word("Second", 60, 0.95), Word("Weak", 100, 0.1)
        }
    }));

    private static OcrTextSpan Word(string text, double y, double confidence) => new OcrTextSpan {
        Text = text, Confidence = confidence, Level = OcrTextSpanLevel.Word,
        CoordinateUnit = OcrCoordinateUnit.Points, Region = new OcrRegion { X = 20, Y = y, Width = 80, Height = 12 }
    };
}
