using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.Tests;

public sealed class PowerPointFinalGeometryLengthTests {
    [Fact]
    public void StrokeAndCellWidthsPreserveExactEmusNullsAndDrawingLimitsInSavedPptx() {
        using var output = new MemoryStream();
        using (var document = PowerPointDocument.Create()) {
            var slide = document.AddSlide();
            var shape = slide.AddRectangle(OfficeLength.Points(10), OfficeLength.Points(10),
                OfficeLength.Points(80), OfficeLength.Points(40), "Exact outline");
            shape.Stroke("123456", OfficeLength.Emus(1));
            shape.Stroke("654321");
            Assert.Equal(OfficeLength.Emus(1), shape.OutlineWidth);
            var inherited = slide.AddRectangle(OfficeLength.Points(100), OfficeLength.Points(10),
                OfficeLength.Points(80), OfficeLength.Points(40), "Inherited width");
            inherited.OutlineWidth = OfficeLength.Points(2);
            inherited.OutlineWidth = null;
            Assert.Null(inherited.OutlineWidth);
            var zero = slide.AddRectangle(OfficeLength.Points(190), OfficeLength.Points(10),
                OfficeLength.Points(80), OfficeLength.Points(40), "Zero width");
            zero.Stroke("ABCDEF", OfficeLength.Zero);
            var connector = slide.AddConnectionShape(A.ShapeTypeValues.StraightConnector1, 1, 2, 3, 4);
            connector.Stroke("123456", OfficeLength.Points(1584));
            connector.Stroke("654321");
            Assert.Equal(OfficeLength.Points(1584), connector.OutlineWidth);
            var table = slide.AddTable(1, 2);
            table.SetCellBorders(PowerPointTableCellBorders.All, "123456", OfficeLength.Emus(1));
            var cell = table.GetCell(0, 0);
            cell.SetBorders(PowerPointTableCellBorders.Left, "654321", OfficeLength.Points(1584));
            cell.SetBorders(PowerPointTableCellBorders.Left, "ABCDEF");
            cell.SetBorders(PowerPointTableCellBorders.Bottom, "654321", OfficeLength.Zero, PowerPointLineDashStyle.Dash);

            string before = slide.SlidePart.Slide.OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.Stroke("FFFFFF", OfficeLength.Emus(20116801)));
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.OutlineWidth = OfficeLength.Emus(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => connector.Stroke("FFFFFF", OfficeLength.Emus(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellBorders(
                PowerPointTableCellBorders.All, "FFFFFF", OfficeLength.Emus(20116801)));
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.SetBorders(
                PowerPointTableCellBorders.All, "FFFFFF", OfficeLength.Emus(-1), PowerPointLineDashStyle.Dash));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellBorders(
                PowerPointTableCellBorders.All, "FFFFFF", OfficeLength.Points(1), (PowerPointLineDashStyle)int.MaxValue));
            Assert.Equal(before, slide.SlidePart.Slide.OuterXml);
            document.Save(output);
        }

        output.Position = 0;
        using (var package = PresentationDocument.Open(output, false)) {
            var slide = package.PresentationPart!.SlideParts.Single().Slide;
            var shapes = slide.Descendants<Shape>().ToArray();
            Assert.Equal(1, shapes.Single(item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name?.Value == "Exact outline")
                .ShapeProperties!.GetFirstChild<A.Outline>()!.Width!.Value);
            Assert.Null(shapes.Single(item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name?.Value == "Inherited width")
                .ShapeProperties!.GetFirstChild<A.Outline>()?.Width);
            Assert.Equal(0, shapes.Single(item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name?.Value == "Zero width")
                .ShapeProperties!.GetFirstChild<A.Outline>()!.Width!.Value);
            Assert.Equal(20116800, slide.Descendants<ConnectionShape>().Single().ShapeProperties!
                .GetFirstChild<A.Outline>()!.Width!.Value);
            var cells = slide.Descendants<A.TableCell>().ToArray();
            Assert.Equal(20116800, cells[0].TableCellProperties!.LeftBorderLineProperties!.Width!.Value);
            Assert.Equal(1, cells[0].TableCellProperties!.TopBorderLineProperties!.Width!.Value);
            Assert.Equal(0, cells[0].TableCellProperties!.BottomBorderLineProperties!.Width!.Value);
            Assert.Equal(A.PresetLineDashValues.Dash, cells[0].TableCellProperties!.BottomBorderLineProperties!
                .GetFirstChild<A.PresetDash>()!.Val!.Value);
            Assert.Equal(1, cells[1].TableCellProperties!.LeftBorderLineProperties!.Width!.Value);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        Assert.Equal(OfficeLength.Emus(1), loaded.Slides[0].Shapes.Single(shape => shape.Name == "Exact outline").OutlineWidth);
    }

    [Fact]
    public void CommentPositionsPreserveSignedNativeLimitsAndRejectBeforeAddingParts() {
        using var output = new MemoryStream();
        using (var document = PowerPointDocument.Create()) {
            var slide = document.AddSlide();
            var author = new PowerPointCommentAuthor("Geometry reviewer", "GR");
            Assert.Throws<ArgumentOutOfRangeException>(() => document.AddClassicComment(slide, author, "Invalid",
                y: OfficeLength.Points(((long)int.MaxValue + 1) / 8D)));
            Assert.Throws<ArgumentOutOfRangeException>(() => document.AddModernComment(slide, author, "Invalid",
                y: OfficeLength.Emus(27273042316901L)));
            Assert.Null(document.OpenXmlDocument.PresentationPart!.CommentAuthorsPart);
            Assert.Null(slide.SlidePart.SlideCommentsPart);
            Assert.Empty(document.OpenXmlDocument.PresentationPart.Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointAuthorsPart>());
            Assert.Empty(document.GetClassicComments(slide));
            Assert.Empty(document.GetModernComments(slide));

            var classic = document.AddClassicComment(slide, author, "Classic",
                OfficeLength.Points(int.MinValue / 8D), OfficeLength.Points(int.MaxValue / 8D));
            var modern = document.AddModernComment(slide, author, "Modern",
                x: OfficeLength.Emus(-27273042329600L), y: OfficeLength.Emus(27273042316900L));
            Assert.Throws<ArgumentOutOfRangeException>(() => classic.X = OfficeLength.Points(((long)int.MinValue - 1) / 8D));
            Assert.Throws<ArgumentOutOfRangeException>(() => modern.Y = OfficeLength.Emus(27273042316901L));
            Assert.Equal(OfficeLength.Points(int.MinValue / 8D), classic.X);
            Assert.Equal(OfficeLength.Emus(27273042316900L), modern.Y);
            var origin = document.AddClassicComment(slide, author, "Origin");
            Assert.Equal(OfficeLength.Zero, origin.X);
            Assert.Equal(OfficeLength.Zero, origin.Y);
            document.Save(output);
        }
        output.Position = 0;
        using (var package = PresentationDocument.Open(output, false)) {
            var slide = package.PresentationPart!.SlideParts.Single();
            var classic = slide.SlideCommentsPart!.CommentList!.Elements<Comment>().First();
            Assert.Equal((long)int.MinValue, classic.Position!.X!.Value);
            Assert.Equal((long)int.MaxValue, classic.Position.Y!.Value);
            var modern = slide.Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>().Single()
                .CommentList!.Elements<P188.Comment>().Single().GetFirstChild<P188.Point2DType>()!;
            Assert.Equal(-27273042329600L, modern.X!.Value);
            Assert.Equal(27273042316900L, modern.Y!.Value);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        Assert.Equal(OfficeLength.Points(int.MaxValue / 8D), loaded.GetClassicComments(loaded.Slides[0])[0].Y);
        Assert.Equal(OfficeLength.Emus(-27273042329600L), loaded.GetModernComments(loaded.Slides[0])[0].X);
        Assert.Equal(OfficeLength.Points(int.MinValue / 8D), loaded.InspectReviewComments().Comments[0].X);
    }

    [Theory]
    [InlineData(PowerPointFileFormat.Pptx)]
    [InlineData(PowerPointFileFormat.Ppt)]
    public void ClassicCommentPositionsUseEighthPointsInPptxAndBinaryPpt(PowerPointFileFormat format) {
        byte[] bytes;
        using (var document = PowerPointDocument.Create()) {
            var slide = document.AddSlide();
            var author = new PowerPointCommentAuthor("Geometry reviewer", "GR");
            var inch = document.AddClassicComment(slide, author, "One inch", x: OfficeLength.Inches(1));
            inch.Y = OfficeLength.Inches(-1);
            var quantized = document.AddClassicComment(slide, author, "Quantized", x: OfficeLength.Emus(1587));
            quantized.Y = OfficeLength.Emus(-1587);
            Assert.Equal(OfficeLength.Emus(1588), quantized.X);
            Assert.Equal(OfficeLength.Emus(-1588), quantized.Y);
            bytes = document.ToBytes(new PowerPointSaveOptions { Format = format });
        }

        if (format == PowerPointFileFormat.Pptx) {
            using var package = PresentationDocument.Open(new MemoryStream(bytes, writable: false), false);
            var comments = package.PresentationPart!.SlideParts.Single().SlideCommentsPart!
                .CommentList!.Elements<Comment>().ToArray();
            Assert.Equal(576L, comments[0].Position!.X!.Value);
            Assert.Equal(-576L, comments[0].Position!.Y!.Value);
            Assert.Equal(1L, comments[1].Position!.X!.Value);
            Assert.Equal(-1L, comments[1].Position!.Y!.Value);
        } else {
            var comments = Assert.Single(LegacyPptDocument.Load(bytes).Slides).Comments;
            Assert.Equal(576, comments[0].X);
            Assert.Equal(-576, comments[0].Y);
            Assert.Equal(1, comments[1].X);
            Assert.Equal(-1, comments[1].Y);
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        var editable = loaded.GetClassicComments(loaded.Slides[0]);
        Assert.Equal(OfficeLength.Inches(1), editable[0].X);
        Assert.Equal(OfficeLength.Inches(-1), editable[0].Y);
        Assert.Equal(OfficeLength.Emus(1588), editable[1].X);
        Assert.Equal(OfficeLength.Emus(-1588), editable[1].Y);
        var inspected = loaded.InspectReviewComments().Comments;
        Assert.Equal(OfficeLength.Inches(1), inspected[0].X);
        Assert.Equal(OfficeLength.Inches(-1), inspected[0].Y);
        Assert.Equal(OfficeLength.Emus(1588), inspected[1].X);
        Assert.Equal(OfficeLength.Emus(-1588), inspected[1].Y);
    }

    [Fact]
    public void CompositionDefaultsAndExplicitRegionsSaveWithoutLosingSingleEmus() {
        using var document = PowerPointDocument.Create();
        PowerPointLayoutBox expected = default;
        var plan = new PowerPointDeckPlan().AddCustom("Geometry", context => {
            var area = context.ContentArea();
            Assert.Equal(OfficeLength.Centimeters(3.65), area.Top);
            Assert.Equal(OfficeLength.Centimeters(1.5), area.Left);
            var columns = context.ContentColumns(2);
            Assert.Equal(OfficeLength.Centimeters(0.65), columns[1].Left - columns[0].Right);
            var rows = context.ContentRows(2);
            Assert.Equal(OfficeLength.Centimeters(0.55), rows[1].Top - rows[0].Bottom);
            var grid = context.ContentGrid(2, 2);
            Assert.Equal(OfficeLength.Centimeters(0.55), grid[0, 1].Left - grid[0, 0].Right);
            Assert.Equal(OfficeLength.Centimeters(0.55), grid[1, 0].Top - grid[0, 0].Bottom);
            Assert.Equal(area, context.ContentArea(top: null, bottomMargin: null, horizontalMargin: null));
            Assert.Equal(area, context.UsePreset(PowerPointCompositionPreset.MetricStory).Content);
            Assert.Equal(area, context.UsePreset(PowerPointCompositionPreset.MetricStory, PowerPointCompositionVariant.Standard).Content);
            var exact = context.ContentArea(OfficeLength.Emus(100001), OfficeLength.Emus(200003), OfficeLength.Emus(300007));
            Assert.Equal(OfficeLength.Emus(100001), exact.Top);
            Assert.Equal(OfficeLength.Emus(300007), exact.Left);
            expected = exact;
            string before = context.Slide.SlidePart.Slide.OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => context.ContentArea(horizontalMargin: OfficeLength.Emus(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => context.ContentColumns(2, OfficeLength.Emus(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => context.ContentGrid(2, 2, OfficeLength.Emus(-1)));
            Assert.Equal(before, context.Slide.SlidePart.Slide.OuterXml);
            context.Slide.AddRectangle(exact.Left, exact.Top, exact.Width, exact.Height, "Exact region");
        });
        var options = PowerPointCompositionOptions.FromBrief(PowerPointDesignBrief.FromBrand("#008C95", "geometry"));
        options.SelectBestAlternative = false;
        document.Compose(plan, options);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using (var package = PresentationDocument.Open(output, false)) {
            var shape = package.PresentationPart!.SlideParts.Single().Slide.Descendants<Shape>()
                .Single(item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name?.Value == "Exact region");
            Assert.Equal(expected.Left.ToEmus(), shape.ShapeProperties!.Transform2D!.Offset!.X!.Value);
            Assert.Equal(expected.Top.ToEmus(), shape.ShapeProperties.Transform2D.Offset.Y!.Value);
            Assert.Equal(expected.Width.ToEmus(), shape.ShapeProperties.Transform2D.Extents!.Cx!.Value);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        var slide = loaded.Slides.Single();
        var region = slide.Shapes.Single(shape => shape.Name == "Exact region");
        Assert.Contains(region, slide.GetShapesInBounds(expected.Left, expected.Top, expected.Width, expected.Height, includePartial: false));
        Assert.Contains(region, slide.GetShapesInBounds(expected, includePartial: true));
        Assert.DoesNotContain(region, slide.GetShapesInBounds(OfficeLength.Emus(1), OfficeLength.Emus(1), OfficeLength.Emus(1), OfficeLength.Emus(1)));
    }

    [Fact]
    public void AccessibilityThresholdsClassifySavedTextAtSingleEmuBoundaries() {
        using var output = new MemoryStream();
        using (var document = PowerPointDocument.Create()) {
            var slide = document.AddSlide();
            slide.BackgroundColor = "FFFFFF";
            var regular = slide.AddTextBox("Regular threshold", OfficeLength.Points(10), OfficeLength.Points(10),
                OfficeLength.Points(200), OfficeLength.Points(40));
            regular.FontSize = OfficeLength.Points(18);
            regular.Color = "888888";
            regular.FillColor = "FFFFFF";
            var bold = slide.AddTextBox("Bold threshold", OfficeLength.Points(10), OfficeLength.Points(70),
                OfficeLength.Points(200), OfficeLength.Points(40));
            bold.FontSize = OfficeLength.Points(14);
            bold.Bold = true;
            bold.Color = "888888";
            bold.FillColor = "FFFFFF";
            document.Save(output);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        Assert.DoesNotContain(loaded.InspectAccessibility().Findings, finding => finding.Code == "accessibility-low-contrast");
        var options = new PowerPointAccessibilityOptions {
            LargeTextThreshold = OfficeLength.Points(18) + OfficeLength.Emus(1),
            LargeBoldTextThreshold = OfficeLength.Points(14) + OfficeLength.Emus(1)
        };
        var findings = loaded.InspectAccessibility(options).Findings
            .Where(finding => finding.Code == "accessibility-low-contrast").ToArray();
        Assert.Equal(2, findings.Length);
        Assert.All(findings, finding => Assert.Equal(4.5D, finding.RequiredValue));
    }

    [Fact]
    public void PolicyLengthsKeepDefaultsCloneAndRejectInvalidValuesWithoutMutation() {
        var accessibility = new PowerPointAccessibilityOptions();
        Assert.Equal(OfficeLength.Points(18), accessibility.LargeTextThreshold);
        Assert.Equal(OfficeLength.Points(14), accessibility.LargeBoldTextThreshold);
        Assert.Throws<ArgumentOutOfRangeException>(() => accessibility.LargeTextThreshold = OfficeLength.Zero);
        Assert.Throws<ArgumentOutOfRangeException>(() => accessibility.LargeBoldTextThreshold = OfficeLength.Emus(-1));
        Assert.Equal(OfficeLength.Points(18), accessibility.LargeTextThreshold);
        Assert.Equal(OfficeLength.Points(14), accessibility.LargeBoldTextThreshold);
        var policy = new PowerPointDeckPreflightOptions();
        Assert.Equal(OfficeLength.Points(10), policy.MinimumReadableFontSize);
        Assert.Equal(OfficeLength.Points(18), policy.DefaultFontSize);
        Assert.Equal(OfficeLength.Points(1), policy.CollisionTolerance);
        Assert.Equal(OfficeLength.Points(36), policy.MaximumDecorativeBleed);
        Assert.Throws<ArgumentOutOfRangeException>(() => policy.MinimumReadableFontSize = OfficeLength.Zero);
        Assert.Throws<ArgumentOutOfRangeException>(() => policy.DefaultFontSize = OfficeLength.Emus(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => policy.CollisionTolerance = OfficeLength.Emus(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => policy.MaximumDecorativeBleed = OfficeLength.Emus(-1));
        policy.MinimumReadableFontSize = OfficeLength.Emus(1);
        policy.DefaultFontSize = OfficeLength.Points(20);
        policy.CollisionTolerance = OfficeLength.Zero;
        policy.MaximumDecorativeBleed = OfficeLength.Zero;
        var clone = policy.Clone();
        Assert.Equal(policy.MinimumReadableFontSize, clone.MinimumReadableFontSize);
        Assert.Equal(policy.DefaultFontSize, clone.DefaultFontSize);
        Assert.Equal(policy.CollisionTolerance, clone.CollisionTolerance);
        Assert.Equal(policy.MaximumDecorativeBleed, clone.MaximumDecorativeBleed);
        clone.DefaultFontSize = OfficeLength.Points(24);
        Assert.Equal(OfficeLength.Points(20), policy.DefaultFontSize);
    }
}
