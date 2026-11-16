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
                // Redaction chooses what to remove from the reader's geometry, which still lets a surplus Q undo the page
                // transformation; Docs/ROADMAP.md tracks reading such pages the way viewers draw them.
                if (shape != ContentShape.Balanced && !(operation == "Redaction" && shape == ContentShape.RepeatedExtraRestores)) {
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

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void TurnedFreeTextKeepsItsCalloutWhereItWasDeclared(int rotation) {
        // No appearance stream, so flattening builds one from /Rotate, /CL, /RD and /C. /Rect encloses the callout and
        // /RD insets the text box to [100 100 260 180].
        byte[] pdf = Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Annots [4 0 R] >>",
            "<< /Type /Annot /Subtype /FreeText /F 4 /Rect [20 20 260 180] /RD [80 80 0 0] /Rotate " + rotation +
                " /Contents (NOTE) /DA (/Helv 10 Tf 0 g) /C [1 0 0] /IT /FreeTextCallout /CL [40 40 100 120] /BS << /W 2 >> >>"
        });

        byte[] flattened = PdfDocument.Load(pdf).Annotations.Flatten().Bytes;

        // The callout starts at user (40, 40), which is (40, 260) on this unrotated page.
        OfficeRasterImage after = Render(flattened);
        Assert.False(Find(after, Ink.Red, 34, 254, 46, 266).IsEmpty, "The callout does not start where /CL declares it.");
        // Its midpoint, user (70, 80), is drawn too: the line runs between the declared points.
        Assert.False(Find(after, Ink.Red, 64, 214, 76, 226).IsEmpty, "The callout does not run between its declared points.");
    }

    [Fact]
    public void AddedTextWithAnExplicitRotationStaysInsideItsRegionOnRotatedPages() {
        var fixture = new Fixture(ContentShape.Balanced, 180, BoxShape.Origin);

        byte[] edited = PdfDocument.Load(fixture.Bytes).Text.Add(new PdfPageRegion(1, 200, 100, 160, 40), "EDITMARK",
            new PdfTextEditOptions { Color = Red, FontSize = 12, RotationDegrees = 0 }).Document.ToBytes();

        // User region [200 100 360 140] is displayed at x 40..200, y 100..140 on the half-turned page.
        Box red = Find(Render(edited), Ink.Red);
        Assert.False(red.IsEmpty);
        Assert.InRange(red.Left, 38, 202);
        Assert.InRange(red.Right, 38, 202);
        Assert.InRange(red.Top, 98, 142);
        Assert.InRange(red.Bottom, 98, 142);
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(0.5D)]
    public void FlatteningKeepsTheOpacityOfAShapeCoveringContent(double opacity) {
        var fixture = new Fixture(ContentShape.Balanced, 0, BoxShape.Origin);
        // The square covers the right half of the blue mark: opaque hides it, translucent tints it.
        byte[] annotated = PdfDocument.Load(fixture.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1, Subtype = "Square", Rectangle = new double[] { 60, 20, 120, 80 }, Opacity = opacity,
            Color = new double[] { 1, 0, 0 }, InteriorColor = new double[] { 1, 0, 0 }, GenerateAppearance = true
        }).Bytes;
        OfficeIMO.OfficeColor covered = Render(annotated).GetPixel(75, 250);
        OfficeIMO.OfficeColor uncovered = Render(annotated).GetPixel(40, 250);

        byte[] flattened = PdfDocument.Load(annotated).Annotations.Flatten().Bytes;

        OfficeRasterImage after = Render(flattened);
        AssertSameColor(covered, after.GetPixel(75, 250));
        AssertSameColor(uncovered, after.GetPixel(40, 250));
        Assert.Equal(opacity < 1D, covered.B > 60);
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

    private static void AssertSameColor(OfficeIMO.OfficeColor expected, OfficeIMO.OfficeColor actual) {
        Assert.InRange(actual.R, expected.R - 3, expected.R + 3);
        Assert.InRange(actual.G, expected.G - 3, expected.G + 3);
        Assert.InRange(actual.B, expected.B - 3, expected.B + 3);
    }

    private static void AssertSameBox(Box expected, Box actual) {
        Assert.False(expected.IsEmpty);
        Assert.InRange(actual.Left, expected.Left - 1, expected.Left + 1);
        Assert.InRange(actual.Top, expected.Top - 1, expected.Top + 1);
        Assert.InRange(actual.Right, expected.Right - 1, expected.Right + 1);
        Assert.InRange(actual.Bottom, expected.Bottom - 1, expected.Bottom + 1);
    }
}
