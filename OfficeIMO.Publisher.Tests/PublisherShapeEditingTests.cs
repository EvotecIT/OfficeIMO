using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherShapeEditingTests {
    [Fact]
    public void Native_rectangle_edits_share_frame_placement_and_reach_svg_and_reopened_pdf_without_mutating_snapshots() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        PublisherPage originalPage = Assert.Single(document.Pages);
        PublisherShape original = Assert.Single(originalPage.Shapes);
        Assert.True(original.CanEdit, original.EditLimitation); Assert.True(original.HasPlacement);
        Assert.Equal((72D, 72D, 144D, 72D), (original.X, original.Y, original.Width, original.Height));
        PublisherTextFrame frame = Assert.Single(originalPage.TextFrames);
        Assert.Equal(frame.Id, original.Id);
        string[] stories = document.TextStories.Select(story => story.Text).ToArray();
        var sourceReport = document.SourceReadReport;
        string oldDrawing = OfficeDrawingSvgExporter.ToSvg(originalPage.CreateDrawing());
        document.SetShapeBounds(original.Id, 36, 48, 120, 84);
        document.SetShapePaint(original.Id, OfficeColor.FromRgb(0, 255, 0), OfficeColor.Black, 2);
        PublisherShape current = Assert.Single(document.Pages[0].Shapes);
        PublisherTextFrame currentFrame = Assert.Single(document.Pages[0].TextFrames);
        Assert.Equal((36D, 48D, 120D, 84D), (current.X, current.Y, current.Width, current.Height));
        Assert.Equal((current.X, current.Y, current.Width, current.Height), (currentFrame.X, currentFrame.Y, currentFrame.Width, currentFrame.Height));
        Assert.Equal(new OfficePoint(36, 48), current.PageTransform.TransformPoint(new OfficePoint(0, 0)));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), Painted(document.Pages[0], current.Id).Shape.FillColor);
        Assert.Contains("#00FF00", document.ToSvg(), StringComparison.OrdinalIgnoreCase);
        byte[] pdf = document.ToPdfBytes();
        Assert.Single(PdfReadDocument.Open(pdf).Pages);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), raster.GetPixel(50, 60));
        Assert.Equal(OfficeColor.White, raster.GetPixel(200, 100));
        Assert.Equal(stories, document.TextStories.Select(story => story.Text));
        Assert.Same(sourceReport, document.SourceReadReport);
        Assert.Equal(oldDrawing, OfficeDrawingSvgExporter.ToSvg(originalPage.CreateDrawing()));
        Assert.Equal(72, original.X); Assert.Equal(144, original.Width);
        OfficeShape detached = current.CreateShape(); detached.FillColor = OfficeColor.Black; detached.Width = 5;
        Assert.Equal(120, current.CreateShape().Width); Assert.Equal(OfficeColor.FromRgb(0, 255, 0), current.CreateShape().FillColor);
        document.SetTextFrameBounds(current.Id, 90, 100, 150, 90);
        PublisherShape movedByFrame = Assert.Single(document.Pages[0].Shapes);
        Assert.Equal((90D, 100D, 150D, 90D), (movedByFrame.X, movedByFrame.Y, movedByFrame.Width, movedByFrame.Height));
        Assert.Equal(150, Painted(document.Pages[0], current.Id).Shape.Width);
        Assert.Equal(36, current.X);
    }

    [Fact]
    public void Primitive_creation_copies_inputs_preserves_order_and_applies_master_artwork_to_every_referring_page() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        PublisherPage oldPage = document.Pages[0], oldMaster = Assert.Single(document.MasterPages);
        int sourceObjects = document.ReadReport.SourceObjectCount, records = document.ReadReport.InspectedRecordCount;
        var input = OfficeShape.Ellipse(40, 30); input.FillColor = OfficeColor.FromRgb(0, 255, 0);
        PublisherShape pageShape = document.AddShape(10, input, 220, 40);
        input.FillColor = OfficeColor.Black; input.Width = 1;
        Assert.Same(pageShape, document.Pages[0].Shapes.Last());
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), pageShape.CreateShape().FillColor); Assert.Equal(40, pageShape.Width);
        OfficeShape masterInput = OfficeShape.RoundedRectangle(40, 24, 4); masterInput.FillColor = OfficeColor.FromRgb(0, 0, 255);
        PublisherShape masterShape = document.AddShape(20, masterInput, 50, 250);
        Assert.Equal(new uint[] { 101, 102, pageShape.Id }, document.Pages[0].Shapes.Select(shape => shape.Id));
        Assert.Equal(new uint[] { 103, masterShape.Id }, document.MasterPages[0].Shapes.Select(shape => shape.Id));
        Assert.Empty(document.Pages[1].Shapes);
        Assert.Equal(new uint[] { 103, masterShape.Id, 101, 102, pageShape.Id }, PaintOrder(document.Pages[0]));
        Assert.All(document.Pages, page => {
            Assert.DoesNotContain(page.Shapes, shape => shape.Id == masterShape.Id);
            Assert.Equal(OfficeColor.FromRgb(0, 0, 255), Painted(page, masterShape.Id).Shape.FillColor);
            Assert.Contains("#0000FF", document.ToSvg(document.Pages.ToList().IndexOf(page)), StringComparison.OrdinalIgnoreCase);
        });
        byte[] pdf = document.ToPdfBytes();
        Assert.Equal(2, PdfReadDocument.Open(pdf).Pages.Count);
        Assert.All(new[] { 1, 2 }, page => Assert.Equal(OfficeColor.FromRgb(0, 0, 255),
            OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[page - 1].ToDrawing()).GetPixel(60, 260)));
        Assert.Equal(new uint[] { 101, 102 }, oldPage.Shapes.Select(shape => shape.Id)); Assert.Single(oldMaster.Shapes);
        Assert.Equal(sourceObjects, document.ReadReport.SourceObjectCount);
        Assert.Equal(sourceObjects + 2, document.ReadReport.CurrentObjectCount);
        Assert.Equal(sourceObjects + 2, document.ReadReport.ProjectedObjectCount); Assert.Equal(records, document.ReadReport.InspectedRecordCount);
        document.SetShapeBounds(masterShape.Id, -5, 240, 50, 30);
        Assert.All(document.Pages, page => Assert.Equal(50, Painted(page, masterShape.Id).Shape.Width));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PAGE_BLEED_CLIPPED"
            && diagnostic.Location == "Contents/object/" + masterShape.Id);
    }

    [Theory]
    [InlineData(0, 30)]
    [InlineData(40, 0)]
    [InlineData(40, 30)]
    public void Authored_line_keeps_logical_zero_extents_and_endpoint_direction_through_resize_and_exports(double width, double height) {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        OfficeShape line = OfficeShape.Line(width, 0, 0, height); line.StrokeColor = OfficeColor.Black; line.StrokeWidth = 2;
        PublisherShape created = document.AddShape(10, line, 30, 110);
        Assert.Equal(width, created.Width); Assert.Equal(height, created.Height);
        Assert.Equal(line.Points, Painted(document.Pages[0], created.Id).Shape.Points);
        document.SetShapeBounds(created.Id, 30, 110, width * 2, height * 2);
        PublisherShape current = document.Pages[0].Shapes.Single(shape => shape.Id == created.Id);
        OfficeShape painted = Painted(document.Pages[0], created.Id).Shape;
        Assert.Equal(width * 2, current.Width); Assert.Equal(height * 2, current.Height);
        Assert.Equal(new OfficePoint(width * 2, 0), painted.Points[0]);
        Assert.Equal(new OfficePoint(0, height * 2), painted.Points[1]);
        Assert.Contains("<line", document.ToSvg());
        Assert.Equal(2, PdfReadDocument.Open(document.ToPdfBytes()).Pages.Count);
        Assert.Throws<ArgumentException>(() => document.SetShapeBounds(created.Id, 30, 110, 0, 0));
    }

    [Theory]
    [InlineData(1, OfficeShapeKind.Rectangle)]
    [InlineData(2, OfficeShapeKind.RoundedRectangle)]
    [InlineData(3, OfficeShapeKind.Ellipse)]
    [InlineData(4, OfficeShapeKind.Polygon)]
    [InlineData(5, OfficeShapeKind.Polygon)]
    [InlineData(6, OfficeShapeKind.Polygon)]
    [InlineData(20, OfficeShapeKind.Line)]
    public void Imported_basic_geometry_resizes_from_the_current_shared_descriptor(int type, OfficeShapeKind kind) {
        byte[] input = PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint> {
            [0x0181] = 0x000000FF, [0x01BF] = 0x00100010, [0x01C0] = 0,
            [0x01FF] = 0x00080008, [0x01CB] = 12700
        }, type);
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherShape original = document.Pages[0].Shapes.Single(shape => shape.Id == 293);
        Assert.True(original.CanEdit, original.EditLimitation);
        OfficeShape before = original.CreateShape();
        document.SetShapeBounds(original.Id, 30, 40, 120, 60);
        OfficeShape current = document.Pages[0].Shapes.Single(shape => shape.Id == original.Id).CreateShape();
        Assert.Equal(kind, current.Kind); Assert.Equal(120, current.Width); Assert.Equal(60, current.Height);
        Assert.Equal(before.FillColor, current.FillColor); Assert.Equal(before.StrokeWidth, current.StrokeWidth);
        Assert.Equal(current.Kind, Painted(document.Pages[0], original.Id).Shape.Kind);
        if (kind == OfficeShapeKind.Polygon) {
            Assert.Equal(before.Points.Count, current.Points.Count);
            for (int index = 0; index < before.Points.Count; index++) {
                Assert.Equal(before.Points[index].X / before.Width * 120, current.Points[index].X, 6);
                Assert.Equal(before.Points[index].Y / before.Height * 60, current.Points[index].Y, 6);
            }
        }
        if (kind == OfficeShapeKind.RoundedRectangle) Assert.Equal(10, current.CornerRadius, 6);
        Assert.Equal(before.Width, original.CreateShape().Width);
    }

    [Fact]
    public void Shapes_and_continuation_frames_share_one_allocator_including_dangling_references() {
        var fixture = PublisherFrameCreationFixture.Load(danglingReference: true);
        PublisherDocument document = fixture.Document;
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        int source = document.ReadReport.SourceObjectCount;
        var red = OfficeShape.Rectangle(20, 20); red.FillColor = OfficeColor.FromRgb(255, 0, 0);
        PublisherShape firstShape = document.AddShape(document.Pages[1].Id, red, 10, 10);
        PublisherTextFrame second = document.AddLinkedTextFrame(first.Id, document.Pages[1].Id, 72, 72, 100, 100);
        PublisherShape secondShape = document.AddShape(document.Pages[1].Id, red, 140, 10);
        PublisherTextFrame third = document.AddLinkedTextFrame(second.Id, document.Pages[1].Id, 200, 72, 100, 100);
        Assert.True(firstShape.Id > fixture.ReservedId);
        Assert.Equal(firstShape.Id + 1, second.Id); Assert.Equal(second.Id + 1, secondShape.Id); Assert.Equal(secondShape.Id + 1, third.Id);
        Assert.Equal(new[] { firstShape.Id, secondShape.Id }, document.Pages[1].Shapes.Select(shape => shape.Id));
        Assert.Equal(new[] { second.Id, third.Id }, document.Pages[1].TextFrames.Select(frame => frame.Id));
        Assert.Equal(source + 4, document.ReadReport.CurrentObjectCount); Assert.Equal(source, document.ReadReport.SourceObjectCount);
    }

    [Fact]
    public void Current_shape_bounds_reflow_referenced_text_and_preserve_the_complete_story() {
        PublisherDocument document = PublisherShapeEditingFixture.Create(withWrapping: true);
        PublisherTextFrame beforeFrame = document.Pages[0].TextFrames.Single(frame => frame.Id == 201);
        string complete = document.TextStories.Single().Text;
        var before = TextRegions(document.Pages[0], 201).Select(text => (text.X, text.Y, text.Width, text.Height)).ToArray();
        Assert.NotEmpty(before);
        document.SetShapeBounds(101, 180, 30, 40, 72);
        var after = TextRegions(document.Pages[0], 201).Select(text => (text.X, text.Y, text.Width, text.Height)).ToArray();
        Assert.NotEqual(before, after);
        Assert.All(TextRegions(document.Pages[0], 201), region => Assert.False(region.X < 220 - .001 && region.X + region.Width > 180 + .001
            && region.Y < 102 - .001 && region.Y + region.Height > 30 + .001));
        Assert.Equal(complete, document.TextStories.Single().Text);
        Assert.Equal(beforeFrame.WrapObjectIds, document.Pages[0].TextFrames.Single(frame => frame.Id == 201).WrapObjectIds);
    }

    [Theory]
    [InlineData("gradient")]
    [InlineData("shadow")]
    [InlineData("perspective")]
    [InlineData("rotation")]
    [InlineData("reflected")]
    [InlineData("hidden")]
    [InlineData("unresolved-color")]
    [InlineData("dash")]
    [InlineData("custom-path")]
    [InlineData("unknown-preset")]
    [InlineData("grouped")]
    [InlineData("no-anchor")]
    public void Unsupported_shape_edits_are_explicit_and_do_not_replace_readable_native_artwork(string restriction) {
        PublisherDocument document = PublisherShapeEditingFixture.Create(restriction);
        PublisherShape shape = document.Pages[0].Shapes.Single(item => item.Id == 101);
        Assert.False(shape.CanEdit); Assert.False(string.IsNullOrEmpty(shape.EditLimitation));
        Assert.Equal(restriction != "no-anchor", shape.HasPlacement);
        Assert.Throws<NotSupportedException>(() => shape.CreateShape());
        string before = document.ToSvg(); var pages = document.Pages; var report = document.ReadReport;
        Assert.Throws<NotSupportedException>(() => document.SetShapeBounds(101, 10, 10, 50, 50));
        Assert.Throws<NotSupportedException>(() => document.SetShapePaint(101, OfficeColor.White, OfficeColor.Black));
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_SHAPE_EDIT_UNSUPPORTED"
            && diagnostic.Location == "Contents/object/101" && diagnostic.LossKind == OfficeConversionLossKind.None);
        if (restriction == "gradient") Assert.NotNull(Painted(document.Pages[0], 101).Shape.FillGradient);
        if (restriction == "unknown-preset") Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_SHAPE_GEOMETRY_APPROXIMATED");
        Assert.NotNull(Painted(document.Pages[0], 102));
    }

    [Fact]
    public void Native_authored_gradient_decorated_line_and_rotated_diamond_remain_inspection_only_with_existing_loss_evidence() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("native-authored.pub"));
        PublisherPage page = document.Pages[0];
        // Native bindings independently identify these three authored objects.
        // In particular, line 295 uses Contents kind 0x20 without an image
        // binding; it must not disappear into the picture inspection model.
        PublisherShape gradient = Assert.Single(page.Shapes, shape => shape.Id == 294);
        PublisherShape line = Assert.Single(page.Shapes, shape => shape.Id == 295);
        PublisherShape diamond = Assert.Single(page.Shapes, shape => shape.Id == 296);
        PublisherShape[] uneditable = new[] { gradient, line, diamond };
        Assert.All(uneditable, shape => { Assert.False(shape.CanEdit); Assert.True(shape.HasPlacement); });
        Assert.DoesNotContain(page.Pictures, picture => picture.Id == line.Id);
        Assert.Equal((72D, 216D, 216D, 0D), (line.X, line.Y, line.Width, line.Height));
        OfficeShape paintedGradient = Painted(page, gradient.Id).Shape;
        Assert.NotNull(paintedGradient.FillGradient);
        Assert.Contains(paintedGradient.FillGradient!.Stops, stop => stop.Color == OfficeColor.FromRgb(255, 0, 0));
        Assert.Contains(paintedGradient.FillGradient.Stops, stop => stop.Color == OfficeColor.FromRgb(0, 0, 255));
        OfficeShape paintedLine = Painted(page, line.Id).Shape;
        Assert.Equal(OfficeShapeKind.Line, paintedLine.Kind);
        Assert.Equal(OfficeLineMarkerKind.Triangle, paintedLine.StrokeStartMarker!.Kind);
        Assert.Equal(OfficeLineMarkerKind.Diamond, paintedLine.StrokeEndMarker!.Kind);
        Assert.Equal(new double[] { 16, 12 }, paintedLine.StrokeDashArray);
        Assert.Equal(OfficeShapeKind.Polygon, Painted(page, diamond.Id).Shape.Kind);
        Assert.Equal(Math.Cos(Math.PI / 6), diamond.PageTransform.M11, 6);
        Assert.Equal(Math.Sin(Math.PI / 6), diamond.PageTransform.M12, 6);
        string before = document.ToSvg();
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_GRADIENT_INTERPOLATION_APPROXIMATED"
            && diagnostic.Location == "Contents/object/294");
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_LINE_MARKER_APPROXIMATED"
            && diagnostic.Location == "Contents/object/295");
        foreach (PublisherShape shape in uneditable) Assert.Throws<NotSupportedException>(() => document.SetShapePaint(shape.Id, OfficeColor.White, OfficeColor.Black));
        Assert.Equal(before, document.ToSvg());
        Assert.Single(PdfReadDocument.Open(document.ToPdfBytes()).Pages);
    }

    [Fact]
    public void Native_line_descriptor_kind_does_not_create_an_unresolved_picture_and_zero_extent_bounds_remain_editable() {
        PublisherDocument document = PublisherShapeEditingFixture.Create("plain-line");
        PublisherShape line = document.Pages[0].Shapes.Single(shape => shape.Id == 101);
        Assert.True(line.CanEdit, line.EditLimitation); Assert.True(line.HasPlacement);
        Assert.Equal(0, line.Height);
        Assert.DoesNotContain(document.Pages[0].Pictures, picture => picture.Id == line.Id);
        document.SetShapeBounds(line.Id, 20, 100, 120, 0);
        OfficeShape current = Painted(document.Pages[0], line.Id).Shape;
        Assert.Equal(OfficeShapeKind.Line, current.Kind); Assert.Equal(120, current.Width); Assert.Equal(0, current.Height);
        Assert.Equal(new[] { new OfficePoint(0, 0), new OfficePoint(120, 0) }, current.Points);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_line_preset_with_an_image_binding_keeps_its_resolved_or_unresolved_picture_identity(bool unresolved) {
        const string fixture = "NativeComparison/Image/native-image.pub";
        PublisherPicture original = Assert.Single(PublisherDocument.Load(PublisherNativeTests.Fixture(fixture)).Pages[0].Pictures);
        var properties = new Dictionary<ushort, uint>();
        if (unresolved) properties.Add(0x4104, 999);
        byte[] input = PublisherDrawingFixture.Mutate(properties, 20, fixture: fixture, objectId: original.Id);
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherPicture picture = Assert.Single(document.Pages[0].Pictures);
        Assert.Equal(original.Id, picture.Id);
        Assert.Equal(unresolved ? 999 : original.ImageId, picture.ImageId);
        Assert.False(picture.CanEdit);
        Assert.DoesNotContain(document.Pages[0].Shapes, shape => shape.Id == picture.Id);
        if (unresolved) Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_IMAGE_REFERENCE_UNRESOLVED");
        else Assert.Contains(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()), element => element is OfficeDrawingImage
            && element.SourceElementIds?.Contains("publisher-object-" + picture.Id) == true);
    }

    internal static OfficeDrawingShape Painted(PublisherPage page, uint id) => PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingShape>().Single(shape => shape.SourceElementIds?.Contains("publisher-object-" + id) == true);
    private static uint[] PaintOrder(PublisherPage page) => PublisherNativeTests.Elements(page.CreateDrawing()).OfType<OfficeDrawingShape>()
        .Where(shape => shape.SourceElementIds?.Count == 1).Select(shape => uint.Parse(shape.SourceElementIds![0].Substring("publisher-object-".Length))).ToArray();
    private static IEnumerable<OfficeDrawingRichText> TextRegions(PublisherPage page, uint id) => PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingRichText>().Where(text => text.SourceElementIds?.Contains("publisher-object-" + id) == true);
}
