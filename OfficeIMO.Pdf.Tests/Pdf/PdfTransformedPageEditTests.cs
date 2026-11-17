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

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void TurnedFreeTextKeepsItsCalloutWhereItWasDeclared(int rotation) {
        // No appearance stream, so flattening builds one from /Rotate, /CL, /RD and /C. /Rect encloses the callout and
        // /RD (left, top, right, bottom) insets the text box to [100 20 260 100].
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
        // The text box border runs along its left edge at user x 100, displayed y 200..280, and nowhere near the box a
        // misread /RD would give ([20 100 180 180], left edge displayed at y 120..200).
        Assert.False(Find(after, Ink.Red, 97, 225, 103, 255).IsEmpty, "The text box is not where /RD places it.");
        Assert.True(Find(after, Ink.Red, 16, 130, 24, 190).IsEmpty, "A text box border is drawn where /RD does not place it.");
    }

    [Fact]
    public void FreeTextWithNoRotateKeepsAnUnturnedAppearanceOnRotatedPages() {
        var fixture = new Fixture(ContentShape.Balanced, 90, BoxShape.Origin);

        byte[] annotated = PdfDocument.Load(fixture.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1, Subtype = "FreeText", Contents = "NOTEMARK", Flags = 4 | 16,
            Rectangle = new double[] { 100, 100, 260, 140 }, GenerateAppearance = true
        }).Bytes;

        PdfAnnotation annotation = Assert.Single(PdfReadDocument.Open(annotated).Pages[0].GetAnnotations());
        Assert.True(annotation.IsNoRotate);
        PdfDictionary freeText = Assert.Single(PdfSyntax.ParseObjects(annotated, null).Map.Values
            .Select(item => item.Value).OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) && subtype is PdfName { Name: "FreeText" }));
        Assert.False(freeText.Items.ContainsKey("Rotate"));
    }

    [Fact]
    public void IsolationLeavesStreamsWithInlineImagesUnchanged() {
        // The image's RGB data contains a Q token; reading it as a surplus restore would corrupt the pixels.
        const string content = "Q\nq 3 0 0 1 10 10 cm\nBI /W 3 /H 1 /BPC 8 /CS /CS0 ID abc Q xyz EI\nQ\n";
        byte[] pdf = Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /ColorSpace << /CS0 /DeviceRGB >> >> /Contents 4 0 R >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream"
        });

        byte[] stamped = PdfStamper.StampText(pdf, "EDITMARK", new PdfTextStampOptions { Color = Red });

        Assert.Contains("ID abc Q xyz EI", PdfTestContentText.ReadDecoded(stamped), StringComparison.Ordinal);
    }

    [Fact]
    public void SurplusRestoresAroundAnInlineImageStayNoOps() {
        // Both restores have nothing to restore: the translation applies to the rectangle and must not reach the edit.
        const string content = "Q 1 0 0 1 100 0 cm BI /W 1 /H 1 /BPC 8 /CS /RGB ID abc EI Q 0 0 1 rg 10 10 20 20 re f\n";
        byte[] pdf = Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Contents 4 0 R >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream"
        });

        byte[] stamped = PdfStamper.StampText(pdf, "EDITMARK", new PdfTextStampOptions { Color = Red });

        IReadOnlyList<PdfViewerPaint> paints = PdfViewerGeometry.Paint(stamped);
        PdfViewerPaint rectangle = Assert.Single(paints, paint => paint.Kind == "path");
        Assert.Equal((110D, 10D, 130D, 30D), (rectangle.A, rectangle.B, rectangle.C, rectangle.D));
        PdfViewerPaint text = Assert.Single(paints, paint => paint.Kind == "text");
        Assert.Equal((36D, 36D), (text.E, text.F));
        Assert.Contains("ID abc EI", PdfTestContentText.ReadDecoded(stamped), StringComparison.Ordinal);
    }

    [Theory]
    // A literal string that runs across a stream boundary: its Q is text, not a restore.
    [InlineData("BT /F1 12 Tf 20 20 Td (A ", "Q B) Tj ET", "Q B) Tj ET")]
    // Data of an inline image whose colour space is unknown can end the image early and pose as a restore, so the
    // depth after it is an estimate and a later restore that may close the saved state is kept.
    [InlineData("q 1 0 0 1 100 0 cm BI /W 3 /H 1 /BPC 8 /CS /Unknown ID ab EI Q cd EI", "Q 0 0 1 rg 10 10 20 20 re f",
        "Q 0 0 1 rg 10 10 20 20 re f")]
    // A filtered image ends wherever an EI looks plausible: this ASCII85 data really ends at ~>, so its Q is a pixel.
    [InlineData("BI /W 8 /H 1 /BPC 8 /CS /G /F /A85 ID !!!!! EI Q !!~> EI", "0 0 1 rg 10 10 20 20 re f", "EI Q !!~> EI")]
    // Nine bytes of image data run into the next stream, so its Q is the last pixel, not a restore.
    [InlineData("BI /W 9 /H 1 /BPC 8 /CS /G ID abcdefgh", "Q EI", "Q EI")]
    public void IsolationReadsPageStreamsAsOneSequenceBeforeBlankingRestores(string first, string second, string kept) {
        byte[] pdf = Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Contents [4 0 R 5 0 R] >>",
            "<< /Length " + first.Length + " >>\nstream\n" + first + "\nendstream",
            "<< /Length " + second.Length + " >>\nstream\n" + second + "\nendstream"
        });

        byte[] stamped = PdfStamper.StampText(pdf, "EDITMARK", new PdfTextStampOptions { Color = Red });

        Assert.Contains(kept, ReadPageContent(stamped), StringComparison.Ordinal);
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

    [Fact]
    public void AddedTextAtAnObliqueAngleStartsBelowTheRegionTopLeft() {
        byte[] pdf = Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>"
        });

        byte[] edited = PdfDocument.Load(pdf).Text.Add(new PdfPageRegion(1, 100, 100, 200, 200), "MMMM",
            new PdfTextEditOptions { Font = PdfStandardFont.Courier, FontSize = 12, RotationDegrees = -60 }).Document.ToBytes();

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(edited).Pages[0].GetTextSpans(), item => item.Text.Contains("MM"));
        Assert.Equal(100D, span.X, 3);
        Assert.Equal(288D, span.Y, 3);
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

    private static string ReadPageContent(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
            dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
        PdfObject contents = PdfObjectLookup.Resolve(objects, page.Items["Contents"])!;
        IEnumerable<PdfObject> entries = contents is PdfArray array ? array.Items : new[] { contents };
        return string.Join("\n", entries.Select(entry => {
            var stream = (PdfStream)PdfObjectLookup.Resolve(objects, entry)!;
            return PdfEncoding.Latin1GetString(OfficeIMO.Pdf.Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
        }));
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
