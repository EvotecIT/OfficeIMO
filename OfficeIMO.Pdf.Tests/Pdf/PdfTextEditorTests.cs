using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextEditorTests {
    [Fact]
    public void InspectAndFindExposeBoundedTextGeometryAndStyle() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Alpha beta alphabet"))
            .ToBytes();
        IReadOnlyList<PdfTextSpan> spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans();
        PdfPageRegion region = RegionAround(spans);

        PdfRegionText inspected = PdfDocument.Open(pdf).Text.Inspect(region);
        IReadOnlyList<PdfTextMatch> contains = PdfDocument.Open(pdf).Text.Find("alpha");
        IReadOnlyList<PdfTextMatch> words = PdfDocument.Open(pdf).Text.Find("alpha", new PdfTextSearchOptions { WholeWords = true });

        Assert.Contains("Alpha beta alphabet", inspected.Text, StringComparison.Ordinal);
        Assert.NotEmpty(inspected.Spans);
        Assert.True(inspected.FontSize > 0D);
        Assert.Equal(2, contains.Count);
        PdfTextMatch wholeWord = Assert.Single(words);
        Assert.Equal("Alpha", wholeWord.Text, ignoreCase: true);
        Assert.True(wholeWord.Width > 0D);
        Assert.True(wholeWord.Height > 0D);
    }

    [Fact]
    public void ReplaceRemovesOnlyTextAndPreservesIntersectingAnnotation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Keep line above"))
            .Paragraph(paragraph => paragraph.Text("Replace this sentence"))
            .Paragraph(paragraph => paragraph.Text("Keep line below"))
            .ToBytes();
        PdfTextMatch sourceMatch = Assert.Single(PdfDocument.Open(source).Text.Find("Replace this sentence", new PdfTextSearchOptions { MatchCase = true }));
        PdfPageRegion region = new PdfPageRegion(1, sourceMatch.X, sourceMatch.Y, sourceMatch.Width, sourceMatch.Height);
        PdfAnnotationEditResult annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Highlight",
            Rectangle = new[] { region.X, region.Y, region.Right, region.Top },
            QuadPoints = new[] { region.X, region.Top, region.Right, region.Top, region.X, region.Y, region.Right, region.Y }
        });

        PdfTextEditResult result = PdfDocument.Open(annotated.Bytes).Text.Replace(region, "Replacement text");
        string text = result.Document.Read.Text();

        Assert.True(result.AffectedCount > 0);
        Assert.DoesNotContain("Replace this sentence", text, StringComparison.Ordinal);
        Assert.Contains("Replacement text", text, StringComparison.Ordinal);
        Assert.Contains("Keep line above", text, StringComparison.Ordinal);
        Assert.Contains("Keep line below", text, StringComparison.Ordinal);
        Assert.Single(result.Document.Read.AnnotationsBySubtype("Highlight"));
    }

    [Fact]
    public void MoveAndAddProduceInspectableTextAtRequestedLocations() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Stationary line above"))
            .Paragraph(paragraph => paragraph.Text("Move me"))
            .Paragraph(paragraph => paragraph.Text("Stationary line below"))
            .ToBytes();
        PdfTextMatch originalMatch = Assert.Single(PdfDocument.Open(source).Text.Find("Move me", new PdfTextSearchOptions { MatchCase = true }));
        PdfPageRegion sourceRegion = new PdfPageRegion(1, originalMatch.X, originalMatch.Y, originalMatch.Width, originalMatch.Height);

        PdfTextEditResult moved = PdfDocument.Open(source).Text.Move(sourceRegion, 80D, -40D);
        IReadOnlyList<PdfTextMatch> movedMatches = moved.Document.Text.Find("Move me", new PdfTextSearchOptions { MatchCase = true });
        PdfTextMatch movedMatch = Assert.Single(movedMatches);
        Assert.InRange(movedMatch.X, originalMatch.X + 79D, originalMatch.X + 81D);
        Assert.InRange(movedMatch.Y, sourceRegion.Y - 41D, sourceRegion.Top - 39D);
        Assert.Contains("Stationary line above", moved.Document.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Stationary line below", moved.Document.Read.Text(), StringComparison.Ordinal);

        var addRegion = new PdfPageRegion(1, 72D, 120D, 200D, 40D);
        PdfTextEditResult added = moved.Document.Text.Add(addRegion, "Added caption", new PdfTextEditOptions {
            Font = PdfStandardFont.CourierBold,
            FontSize = 14D,
            Color = PdfColor.FromRgb(10, 80, 160)
        });
        PdfTextMatch addedMatch = Assert.Single(added.Document.Text.Find("Added caption", new PdfTextSearchOptions { MatchCase = true }));
        Assert.InRange(addedMatch.X, 71.9D, 72.1D);
        Assert.Equal(PdfStandardFont.CourierBold, addedMatch.SuggestedFont);
    }

    [Fact]
    public void ReplaceAllPreservesUnmatchedTextInTheSameSourceSpan() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Keep heading"))
            .Paragraph(paragraph => paragraph.Text("cat cat dog"))
            .Paragraph(paragraph => paragraph.Text("Keep footer"))
            .ToBytes();

        PdfTextEditResult result = PdfDocument.Open(source).Text.ReplaceAll(
            "cat",
            "fox",
            new PdfTextSearchOptions { MatchCase = true, WholeWords = true });
        string text = result.Document.Read.Text();

        Assert.Equal(2, result.AffectedCount);
        Assert.Contains("fox fox dog", text, StringComparison.Ordinal);
        Assert.DoesNotContain("cat", text, StringComparison.Ordinal);
        Assert.Contains("Keep heading", text, StringComparison.Ordinal);
        Assert.Contains("Keep footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SearchOptionsSnapshotPagesAndRejectInvalidSelection() {
        int[] pages = { 1 };
        var options = new PdfTextSearchOptions { PageNumbers = pages };
        pages[0] = 2;

        Assert.Equal(1, Assert.Single(options.PageNumbers!));
        Assert.Throws<ArgumentOutOfRangeException>(() => options.PageNumbers = new[] { 0 });
        Assert.Throws<ArgumentException>(() => options.PageNumbers = new[] { 1, 1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageRegion(0, 0D, 0D, 10D, 10D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextEditOptions { FontSize = 0D });
    }

    [Fact]
    public void PublicRedactionStillPaintsAndRemovesIntersectingAnnotations() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Secret marker"))
            .ToBytes();
        IReadOnlyList<PdfTextSpan> spans = PdfReadDocument.Open(source).Pages[0].GetTextSpans();
        PdfPageRegion region = RegionAround(spans);
        byte[] annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Highlight",
            Rectangle = new[] { region.X, region.Y, region.Right, region.Top },
            QuadPoints = new[] { region.X, region.Top, region.Right, region.Top, region.X, region.Y, region.Right, region.Y }
        }).Bytes;

        PdfDocument redacted = PdfDocument.Open(annotated).Redactions.Apply(new[] { region.ToRedactionAreaForTest() });

        Assert.DoesNotContain("Secret marker", redacted.Read.Text(), StringComparison.Ordinal);
        Assert.Empty(redacted.Read.AnnotationsBySubtype("Highlight"));
    }

    private static PdfPageRegion RegionAround(IReadOnlyList<PdfTextSpan> spans) {
        double left = spans.Min(static span => span.X);
        double right = spans.Max(static span => span.X + Math.Max(1D, Math.Abs(span.Advance)));
        double bottom = spans.Min(static span => span.Y - span.FontSize * 0.3D);
        double top = spans.Max(static span => span.Y + span.FontSize * 0.9D);
        return new PdfPageRegion(1, left - 0.5D, bottom, right - left + 1D, top - bottom);
    }
}

internal static class PdfTextEditorTestExtensions {
    internal static PdfRedactionArea ToRedactionAreaForTest(this PdfPageRegion region) =>
        new PdfRedactionArea(region.PageNumber, region.X, region.Y, region.Width, region.Height);
}
