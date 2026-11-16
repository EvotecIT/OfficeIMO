using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using static OfficeIMO.Tests.Pdf.PdfTransformedPageEditFixtures;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Edits on pages whose existing content leaves the graphics state changed, or whose geometry is rotated or offset,
/// must land where requested, read upright on the displayed page and leave the original content where it was.
/// </summary>
public class PdfTransformedPageEditTests {
    private static readonly OfficeIMO.OfficeColor Red = OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 0);

    public static TheoryData<string, ContentShape, int, BoxShape> IsolationCases() {
        var data = new TheoryData<string, ContentShape, int, BoxShape>();
        foreach (string operation in PdfTransformedPageEditOperations.AppendingOperations)
            foreach (ContentShape shape in Enum.GetValues(typeof(ContentShape)))
                if (shape != ContentShape.Balanced) {
                    data.Add(operation, shape, 0, BoxShape.Origin);
                    data.Add(operation, shape, 90, BoxShape.OffsetMediaBox);
                }
        return data;
    }

    /// <summary>
    /// Compares every painting operator of the edited page with the same edit on a page whose content is balanced. Q
    /// operators that restore more states than were saved are ignored, as viewers do, so a surplus Q in the original
    /// content cannot reach the states the edit relies on.
    /// </summary>
    [Theory]
    [MemberData(nameof(IsolationCases))]
    public void EditsDrawIdenticallyWhateverStateTheOriginalContentLeaves(string operation, ContentShape shape, int rotation, BoxShape box) {
        var control = new Fixture(ContentShape.Balanced, rotation, box);
        var hostile = new Fixture(shape, rotation, box);

        IReadOnlyList<PdfViewerPaint> expected = PdfViewerGeometry.Paint(PdfTransformedPageEditOperations.Apply(operation, control));
        IReadOnlyList<PdfViewerPaint> actual = PdfViewerGeometry.Paint(PdfTransformedPageEditOperations.Apply(operation, hostile));

        Assert.Equal(expected, actual);
    }

    public static TheoryData<string, int, BoxShape> PlacementCases() {
        var data = new TheoryData<string, int, BoxShape>();
        foreach (string operation in PdfTransformedPageEditOperations.PlacedOperations)
            foreach (int rotation in new[] { 0, 90, 180, 270 })
                foreach (BoxShape box in Enum.GetValues(typeof(BoxShape)))
                    data.Add(operation, rotation, box);
        return data;
    }

    /// <summary>Added content sits at the requested displayed-page position, upright, and the original marks stay put.</summary>
    [Theory]
    [MemberData(nameof(PlacementCases))]
    public void EditsOnRotatedAndOffsetPagesLandUprightWhereRequested(string operation, int rotation, BoxShape box) {
        var fixture = new Fixture(ContentShape.Balanced, rotation, box);
        OfficeRasterImage before = Render(fixture.Bytes);

        byte[] edited = PdfTransformedPageEditOperations.Apply(operation, fixture);
        OfficeRasterImage after = Render(edited);

        // Imposition moves the whole page onto a sheet; its check compares the marks' relative placement instead.
        if (operation != "NUp") AssertOriginalMarksUnchanged(before, after);
        string? problem = PdfTransformedPageEditOperations.CheckPlacement(operation, fixture, after, edited);
        Assert.True(problem is null, problem);
    }

    [Theory]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void FreeTextFlattenedOnRotatedPageKeepsItsUprightAppearance(int rotation) {
        var fixture = new Fixture(ContentShape.Balanced, rotation, BoxShape.Origin);
        var corner1 = fixture.ToUser(60, 120);
        var corner2 = fixture.ToUser(260, 160);
        byte[] annotated = PdfDocument.Load(fixture.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1, Subtype = "FreeText", Contents = "NOTEMARK",
            Rectangle = new[] { Math.Min(corner1.X, corner2.X), Math.Min(corner1.Y, corner2.Y), Math.Max(corner1.X, corner2.X), Math.Max(corner1.Y, corner2.Y) },
            Color = new double[] { 1, 0, 0 }, GenerateAppearance = true
        }).Bytes;
        Box beforeFlatten = Find(Render(annotated), Ink.Dark, 55, 115, 265, 165);

        byte[] flattened = PdfDocument.Load(annotated).Annotations.Flatten().Bytes;

        Assert.False(beforeFlatten.IsEmpty);
        Box afterFlatten = Find(Render(flattened), Ink.Dark, 55, 115, 265, 165);
        Assert.Equal((beforeFlatten.Left, beforeFlatten.Top, beforeFlatten.Right, beforeFlatten.Bottom),
            (afterFlatten.Left, afterFlatten.Top, afterFlatten.Right, afterFlatten.Bottom));
        Assert.True(afterFlatten.PixelWidth > afterFlatten.PixelHeight, $"Flattened note is not horizontal: {afterFlatten}");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(flattened).Pages[0].GetTextSpans(), item => item.Text.Contains("NOTE"));
        Assert.Equal(0D, DisplayAngle(span, rotation), 3);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void FlatteningAnnotationWithTurnedAppearanceMatrixKeepsItsPlacement(int matrixRotation) {
        byte[] pdf = PdfTransformedPageEditOperations.BuildPageWithTurnedAppearance(matrixRotation, widget: false);
        Box before = Find(Render(pdf), Ink.Red);
        Box beforeStrip = Find(Render(pdf), Ink.Green);

        byte[] flattened = PdfDocument.Load(pdf).Annotations.Flatten().Bytes;

        OfficeRasterImage after = Render(flattened);
        AssertSameBox(before, Find(after, Ink.Red));
        AssertSameBox(beforeStrip, Find(after, Ink.Green));
        Assert.Empty(PdfReadDocument.Open(flattened).Pages[0].GetAnnotations());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void FlatteningFieldWithTurnedAppearanceMatrixKeepsItsPlacement(int matrixRotation) {
        byte[] pdf = PdfTransformedPageEditOperations.BuildPageWithTurnedAppearance(matrixRotation, widget: true);
        OfficeRasterImage rendered = Render(pdf);
        Box before = Find(rendered, Ink.Red);
        Box beforeStrip = Find(rendered, Ink.Green);

        byte[] flattened = PdfDocument.Load(pdf).Forms.Flatten().ToBytes();

        OfficeRasterImage after = Render(flattened);
        AssertSameBox(before, Find(after, Ink.Red));
        AssertSameBox(beforeStrip, Find(after, Ink.Green));
        Assert.Empty(PdfReadDocument.Open(flattened).Pages[0].GetAnnotations());
    }

    [Fact]
    public void StampAfterRotatingPageTwiceReadsUprightAndKeepsMarkupAndTheMediaBox() {
        var fixture = new Fixture(ContentShape.Balanced, 0, BoxShape.OffsetMediaBox);
        byte[] annotated = PdfDocument.Load(fixture.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1, Subtype = "Square", Rectangle = new double[] { 300, 300, 340, 330 },
            Color = new double[] { 1, 0, 0 }, GenerateAppearance = true
        }).Bytes;
        byte[] rotated = PdfPageEditor.RotatePagesByWithReadOptions(annotated, 90, null, 1);
        rotated = PdfPageEditor.RotatePagesByWithReadOptions(rotated, 90, null, 1);

        byte[] stamped = PdfStamper.StampText(rotated, "EDITMARK", new PdfTextStampOptions { Color = Red, FontSize = OfficeIMO.OfficeLength.Points(14) });

        PdfReadPage page = PdfReadDocument.Open(stamped).Pages[0];
        Assert.Equal(180, page.GetRotationDegrees());
        Assert.Equal("Square", Assert.Single(page.GetAnnotations()).Subtype);
        PdfPageGeometry geometry = page.GetGeometry();
        Assert.Equal((100D, 200D, 500D, 500D), (geometry.MediaBox!.Left, geometry.MediaBox.Bottom, geometry.MediaBox.Right, geometry.MediaBox.Top));
        Box red = Find(Render(stamped), Ink.Red);
        Assert.InRange(red.Left, 34, 40);
        Assert.InRange(red.Bottom, PageHeight - 40, PageHeight - 34);
        Assert.Equal(0D, DisplayAngle(Assert.Single(page.GetTextSpans(), span => span.Text.Contains("EDIT")), 180), 3);
    }

    private static void AssertOriginalMarksUnchanged(OfficeRasterImage before, OfficeRasterImage after) {
        foreach (Ink ink in new[] { Ink.Blue, Ink.Green }) {
            Box original = Find(before, ink);
            Box kept = Find(after, ink, original.Left - 2, original.Top - 2, original.Right + 2, original.Bottom + 2);
            Assert.False(original.IsEmpty);
            // Added content may overlap a mark; a moved or mirrored mark leaves its place empty.
            Assert.True(kept.Count * 2 >= original.Count, $"{ink} mark moved: {original} became {kept}");
        }
    }

    private static void AssertSameBox(Box expected, Box actual) {
        Assert.False(expected.IsEmpty);
        Assert.InRange(actual.Left, expected.Left - 1, expected.Left + 1);
        Assert.InRange(actual.Top, expected.Top - 1, expected.Top + 1);
        Assert.InRange(actual.Right, expected.Right - 1, expected.Right + 1);
        Assert.InRange(actual.Bottom, expected.Bottom - 1, expected.Bottom + 1);
    }
}
