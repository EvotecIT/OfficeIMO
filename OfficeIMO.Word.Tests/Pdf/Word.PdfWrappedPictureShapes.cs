using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

/// <summary>
/// Expected positions come from Word's own PDF export of the same pictures: a 144pt picture at the
/// left margin with 9pt wrap distances, beside a long paragraph.
/// </summary>
public sealed class PdfWrappedPictureShapeTests {
    private const double Margin = 72D, Size = 144D, Distance = 9D;
    // Word's polygons from pictures it wrapped itself, in its 21600 by 21600 space.
    private static readonly long[] Rectangle = { 0, 0, 0, 21375, 21375, 21375, 21375, 0 };
    private static readonly long[] Triangle = { 0, 0, 0, 21375, 21375, 21375, 21375, 21150, 450, 0 };
    private static readonly long[] OpenToTheRight = { 0, 0, 0, 21375, 21375, 21375, 21375, 14850, 6750, 14400, 6750, 7200, 21375, 6525, 21375, 0 };
    // Polygons Word saved for a 450 by 200pt picture: a wedge that is full width at the top, and a
    // full-width bar over a narrow stem. Their first rows leave no room for a word beside them.
    private static readonly long[] Wedge = { 0, 0, 0, 21384, 324, 21384, 405, 21384, 21546, 162, 21546, 0 };
    private static readonly long[] BarOverStem = { 0, 0, 0, 21384, 5508, 21384, 5508, 10368, 21546, 8748, 21546, 0 };
    // The polygon Word saved for the triangle with a 12pt glow, which encloses the glow.
    private static readonly long[] GlowTriangle = { -1350, -1800, -1800, -1575, -1800, 21375, -1350, 23175, 22725, 23175, 23175, 20250, 23175, 20025, 1575, -1800 };

    // Word clears a square-wrapped picture's effect extent as well as its wrap distance.
    [Theory]
    [InlineData(0L, 225D)]
    [InlineData(152400L, 237D)]
    public void SaveAsPdf_SquareWrapClearsTheEffectExtent(long effect, double expectedLeft) {
        IReadOnlyList<double> rows = RowsBeside(AddPicture(wrap: null, effect));
        Assert.True(rows.Count >= 8, "The paragraph wraps beside the picture.");
        Assert.All(rows, left => Assert.InRange(left, expectedLeft - .5D, expectedLeft + .5D));
    }

    // Each row follows the polygon's right edge over the row widened by the wrap distances, so the
    // rows step along the sloped side and end where Word ends them, 9pt past the polygon.
    [Theory]
    [InlineData("tight")]
    [InlineData("through")]
    public void SaveAsPdf_TightWrapFollowsTheWrapPolygon(string kind) {
        IReadOnlyList<double> rows = RowsBeside(AddPicture(kind, polygon: Triangle));
        Assert.True(rows.Count >= 8, "The paragraph wraps beside the picture.");
        Assert.True(rows[0] < 140D, $"The first row starts near the polygon's top corner, not beside the square frame: {rows[0]}.");
        for (int index = 1; index < rows.Count; index++) Assert.True(rows[index] >= rows[index - 1] - .01D);
        Assert.InRange(rows[rows.Count - 1], 223.1D, 224.1D);
    }

    // Word lets tight-wrapped text into a bay of the polygon that opens towards the text.
    [Fact]
    public void SaveAsPdf_TightWrapEntersAnOpenBayOfThePolygon() {
        IReadOnlyList<double> rows = RowsBeside(AddPicture("tight", polygon: OpenToTheRight));
        Assert.Contains(rows, left => Math.Abs(left - 126D) < .5D);
        Assert.Contains(rows, left => Math.Abs(left - 223.6D) < .5D);
    }

    // When the first rows leave no room, Word resumes at the first row where a word fits beside the
    // polygon, not below the whole picture.
    [Theory]
    [InlineData("wedge")]
    [InlineData("bar")]
    public void SaveAsPdf_TightWrapResumesWhereThePolygonLeavesRoom(string shape) {
        IReadOnlyList<double> rows = RowsBeside(AddPicture("tight", polygon: shape == "wedge" ? Wedge : BarOverStem,
            width: 450D, height: 200D, word: "longerword"));
        Assert.True(rows.Count >= 5, $"Rows resume beside the picture: {rows.Count}.");
        if (shape == "wedge") {
            Assert.True(rows[0] > 400D, $"The first row starts in the narrow room at the right: {rows[0]}.");
            for (int index = 1; index < rows.Count; index++) Assert.True(rows[index] < rows[index - 1]);
        } else {
            // The first row fits under the bar's sloping underside at the right, then rows follow the stem.
            Assert.True(rows[0] > 400D, $"The first row starts under the sloping bar: {rows[0]}.");
            Assert.All(rows.Skip(1), left => Assert.InRange(left, 195.5D, 197D));
        }
    }

    // A first word that has no room beside a square picture moves below it instead of running past the margin.
    [Fact]
    public void SaveAsPdf_FirstWordWithNoRoomBesideSquarePictureMovesBelowIt() {
        Assert.Empty(RowsBeside(AddPicture(null, width: 450D, height: 200D, word: "longerword")));
    }

    // The polygon already encloses the effects, so the effect extent is not added again.
    [Fact]
    public void SaveAsPdf_TightWrapUsesThePolygonThatEnclosesEffects() {
        IReadOnlyList<double> rows = RowsBeside(AddPicture("tight", effect: 152400L, polygon: GlowTriangle));
        Assert.InRange(rows.Max(), 235.1D, 236.1D);
    }

    // Word paints a wrapped picture by z-order even when it carries behindDoc="1", as Word itself saves
    // for tight wrapping and for a picture switched from behind-text. Only a picture without wrapping goes behind.
    [Theory]
    [InlineData("square")]
    [InlineData("tight")]
    public void SaveAsPdf_WrappedPictureWithBehindDocPaintsByZOrder(string kind) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Text");
        AddPagePicture(paragraph, OfficeColor.Red, WordImageTextWrapping.InFrontOfText, 100D, 200D, 1U);
        DW.Anchor wrapped = AddPagePicture(paragraph, OfficeColor.Parse("00FF00"), WordImageTextWrapping.Square, 150D, 250D, 2U);
        wrapped.BehindDoc = true;
        if (kind == "tight") SetPolygonWrap(wrapped, "tight", Rectangle);

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        var pdf = PdfCore.PdfDocument.Load(new MemoryStream(bytes));
        byte[] rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(rendered, out var bitmap));
        byte[] pixels = bitmap!.GetPixels();
        int offset = (275 * bitmap.Width + 175) * 4;
        Assert.Equal("0,255,0", pixels[offset] + "," + pixels[offset + 1] + "," + pixels[offset + 2]);
    }

    private static WordDocument AddPicture(string? wrap, long effect = 0L, long[]? polygon = null,
        double width = Size, double height = Size, string word = "word") {
        WordDocument document = WordDocument.Create();
        string text = string.Join(" ", Enumerable.Range(1, 120).Select(index => word + index));
        WordParagraph paragraph = document.AddParagraph(text);
        using var stream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 20, OfficeColor.Parse("1E5AC8")), OfficeImageExportFormat.Png));
        WordImage image = paragraph.AddImage(stream, "shape.png", OfficeLength.Points(width), OfficeLength.Points(height),
            options: new WordImageOptions { Wrapping = WordImageTextWrapping.Square });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Column;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.HorizontalPositionOffset = OfficeLength.Points(0);
        image.VerticalPositionOffset = OfficeLength.Points(0);
        DW.Anchor anchor = image._Image!.Anchor!;
        anchor.DistanceFromLeft = anchor.DistanceFromRight = anchor.DistanceFromTop = anchor.DistanceFromBottom = (uint)(Distance * 12700D);
        DW.EffectExtent extent = anchor.GetFirstChild<DW.EffectExtent>()!;
        extent.LeftEdge = extent.TopEdge = extent.RightEdge = extent.BottomEdge = effect;
        if (wrap != null) SetPolygonWrap(anchor, wrap, polygon!);
        return document;
    }

    private static void SetPolygonWrap(DW.Anchor anchor, string kind, long[] coordinates) {
        var polygon = new DW.WrapPolygon { Edited = false };
        polygon.Append(new DW.StartPoint { X = coordinates[0], Y = coordinates[1] });
        for (int index = 2; index < coordinates.Length; index += 2)
            polygon.Append(new DW.LineTo { X = coordinates[index], Y = coordinates[index + 1] });
        polygon.Append(new DW.LineTo { X = coordinates[0], Y = coordinates[1] });
        DocumentFormat.OpenXml.OpenXmlElement wrap = kind == "tight"
            ? new DW.WrapTight(polygon) { WrapText = DW.WrapTextValues.BothSides }
            : new DW.WrapThrough(polygon) { WrapText = DW.WrapTextValues.BothSides };
        anchor.GetFirstChild<DW.WrapSquare>()!.InsertAfterSelf(wrap);
        anchor.RemoveAllChildren<DW.WrapSquare>();
    }

    private static DW.Anchor AddPagePicture(WordParagraph paragraph, OfficeColor color, WordImageTextWrapping wrapping,
        double left, double top, uint zOrder) {
        using var stream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 20, color), OfficeImageExportFormat.Png));
        WordImage image = paragraph.AddImage(stream, "fill.png", OfficeLength.Points(100), OfficeLength.Points(100),
            options: new WordImageOptions { Wrapping = wrapping });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = OfficeLength.Points(left);
        image.VerticalPositionOffset = OfficeLength.Points(top);
        image._Image!.Anchor!.RelativeHeight = zOrder;
        return image._Image.Anchor;
    }

    /// <summary>The left edge of each text row on the first page that starts beside the picture.</summary>
    private static IReadOnlyList<double> RowsBeside(WordDocument document) {
        using (document) {
            var evidence = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
            Assert.DoesNotContain(evidence.Warnings, warning => warning.Code.Contains("anchored-image", StringComparison.Ordinal));
            byte[] bytes = evidence.Value.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
            using var pdf = PdfPigDocument.Open(bytes);
            return pdf.GetPage(1).GetWords()
                .GroupBy(word => Math.Round(word.BoundingBox.Bottom))
                .OrderByDescending(row => row.Key)
                .Select(row => row.Min(word => word.BoundingBox.Left))
                .Where(left => left > Margin + 5D)
                .ToList();
        }
    }
}
