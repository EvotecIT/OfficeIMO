using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void PageBorder_RendersAsPageDecorationWithOpacityAndDashStyle() {
        byte[] bytes = PdfDocument.Create()
            .PageBorder(PdfColor.FromRgb(30, 64, 175), width: 2, inset: 30, opacity: 0.4, dashStyle: OfficeStrokeDashStyle.Dash)
            .H1("Page border proof")
            .Paragraph(p => p.Text("Body content stays inside the reusable page frame."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        int borderDraw = stream.IndexOf("30 30 552 732 re", StringComparison.Ordinal);
        int firstTextBlock = stream.IndexOf("BT", StringComparison.Ordinal);

        Assert.Contains("/ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("/CA 0.4", raw, StringComparison.Ordinal);
        Assert.Contains("[6 3] 0 d", stream, StringComparison.Ordinal);
        Assert.True(borderDraw >= 0, "Expected the page border rectangle to be drawn.");
        Assert.True(firstTextBlock > borderDraw, "Expected body text to be emitted after the page border decoration.");
    }

    [Fact]
    public void PageBorder_CanBeScopedAndClearedPerComposedPage() {
        byte[] bytes = PdfDocument.Create()
            .PageBorder(inset: 36)
            .Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Framed page"))));
            })
            .Page(page => {
                page.PageBorder((PdfPageBorder?)null);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Clean page"))));
            })
            .ToBytes();

        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));

        Assert.Contains("36 36 540 720 re", firstPageStream, StringComparison.Ordinal);
        Assert.DoesNotContain("36 36 540 720 re", secondPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void PageBorder_ValidatesAndClonesOptions() {
        var options = new PdfOptions {
            PageBorder = new PdfPageBorder {
                Color = PdfColor.FromRgb(1, 2, 3),
                Width = 2,
                Inset = 24,
                Opacity = 0.75,
                DashStyle = OfficeStrokeDashStyle.Dot
            }
        };

        PdfPageBorder snapshot = options.PageBorder!;
        snapshot.Width = 6;

        Assert.Equal(2, options.PageBorder!.Width);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Width = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Inset = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Opacity = double.NaN });

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageBorder = new PdfPageBorder {
                        Inset = 400
                    }
                })
                .Paragraph(p => p.Text("Invalid border frame"))
                .ToBytes());

        Assert.Contains("PDF page border inset must leave a positive border rectangle.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundImage_RendersBehindContentWithOpacityAndFit() {
        byte[] bytes = PdfDocument.Create()
            .BackgroundImage(CreateMinimalRgbPng(), OfficeImageFit.Stretch, opacity: 0.3)
            .H1("Background image proof")
            .Paragraph(p => p.Text("Body content stays above the fitted page background image."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        int imageDraw = stream.IndexOf("/Im", StringComparison.Ordinal);
        int firstTextBlock = stream.IndexOf("BT", StringComparison.Ordinal);

        Assert.Contains("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.Contains("/ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.3", raw, StringComparison.Ordinal);
        Assert.Matches(@"612 0 -?0 792 0 0 cm", stream);
        Assert.True(imageDraw >= 0, "Expected the page background image XObject to be drawn.");
        Assert.True(firstTextBlock > imageDraw, "Expected body text to be emitted after the page background image.");
    }

    [Fact]
    public void BackgroundImage_FallsBackToPageBoxWhenDimensionsAreUnknown() {
        byte[] legacyJpegLikeImage = { 0xFF, 0xD8, 0xFF, 0xD9 };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .BackgroundImage(legacyJpegLikeImage, OfficeImageFit.Cover, opacity: 0.3)
            .Paragraph(p => p.Text("Background image proof"))
            .ToBytes();

        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));

        Assert.Contains("/Im", stream, StringComparison.Ordinal);
        Assert.Matches(@"240 0 -?0 180 0 0 cm", stream);
    }

    [Fact]
    public void BackgroundImage_CanBeScopedAndClearedPerComposedPage() {
        byte[] image = CreateMinimalRgbPng();
        byte[] bytes = PdfDocument.Create()
            .BackgroundImage(image, OfficeImageFit.Stretch, opacity: 0.2)
            .Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Background page"))));
            })
            .Page(page => {
                page.BackgroundImage((PdfPageBackgroundImage?)null);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Clean page"))));
            })
            .ToBytes();

        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));

        Assert.Contains("/Im", firstPageStream, StringComparison.Ordinal);
        Assert.DoesNotContain("/Im", secondPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundImage_ValidatesAndClonesOptions() {
        byte[] image = CreateMinimalRgbPng();
        var options = new PdfOptions {
            PageBackgroundImage = new PdfPageBackgroundImage(image) {
                Fit = OfficeImageFit.Contain,
                Opacity = 0.25
            }
        };

        PdfPageBackgroundImage snapshot = options.PageBackgroundImage!;
        snapshot.Opacity = 0.9;

        Assert.Equal(0.25, options.PageBackgroundImage!.Opacity);
        Assert.Throws<ArgumentException>(() => new PdfPageBackgroundImage(Array.Empty<byte>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBackgroundImage(image) { Opacity = 1.5 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBackgroundImage(image) { Fit = (OfficeImageFit)99 });
    }

    [Fact]
    public void BackgroundShape_RendersBehindContentWithOpacityAndVectorGeometry() {
        var shape = OfficeShape.RoundedRectangle(540, 86, 18);
        shape.FillColor = PdfColor.FromRgb(234, 244, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(96, 165, 250).ToOfficeColor();
        shape.StrokeWidth = 1.25;
        shape.FillOpacity = 0.34;
        shape.StrokeOpacity = 0.6;

        byte[] bytes = PdfDocument.Create()
            .BackgroundShape(new PdfPageBackgroundShape(shape, 36, 640))
            .H1("Background shape proof")
            .Paragraph(p => p.Text("Body content stays above reusable vector page decoration."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        int shapePath = stream.IndexOf("54 640", StringComparison.Ordinal);
        int firstTextBlock = stream.IndexOf("BT", StringComparison.Ordinal);

        Assert.Contains("/ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.34", raw, StringComparison.Ordinal);
        Assert.Contains("/CA 0.6", raw, StringComparison.Ordinal);
        Assert.True(shapePath >= 0, "Expected the rounded background shape path to be drawn.");
        Assert.True(firstTextBlock > shapePath, "Expected body text to be emitted after the background shape.");
    }

    [Fact]
    public void BackgroundShape_CanBeScopedAndClearedPerComposedPage() {
        byte[] bytes = PdfDocument.Create()
            .BackgroundRectangle(36, 640, 540, 86, PdfColor.FromRgb(238, 242, 255))
            .Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Decorated page"))));
            })
            .Page(page => {
                page.ClearBackgroundShapes();
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Clean page"))));
            })
            .ToBytes();

        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));

        Assert.Contains("36 640 540 86 re", firstPageStream, StringComparison.Ordinal);
        Assert.DoesNotContain("36 640 540 86 re", secondPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundBands_ComputePageAnchoredGeometry() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 400,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 40,
                MarginBottom = 40
            })
            .BackgroundTopBand(50, PdfColor.FromRgb(238, 242, 255), insetX: 12, offsetY: 8, stroke: PdfColor.FromRgb(96, 165, 250), strokeWidth: 0.8, fillOpacity: 0.42, strokeOpacity: 0.7, fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.LightBlue, OfficeColor.WhiteSmoke))
            .BackgroundBottomBand(30, PdfColor.FromRgb(240, 253, 244), insetX: 20, offsetY: 10)
            .BackgroundLeftBand(18, PdfColor.FromRgb(254, 249, 195), insetY: 24, offsetX: 6)
            .BackgroundRightBand(22, PdfColor.FromRgb(255, 237, 213), insetY: 30, offsetX: 9)
            .Paragraph(p => p.Text("Anchored bands"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        int topBand = stream.IndexOf("12 342 276 50 re", StringComparison.Ordinal);
        int bottomBand = stream.IndexOf("20 10 260 30 re", StringComparison.Ordinal);
        int leftBand = stream.IndexOf("6 24 18 352 re", StringComparison.Ordinal);
        int rightBand = stream.IndexOf("269 30 22 340 re", StringComparison.Ordinal);
        int firstTextBlock = stream.IndexOf("BT", StringComparison.Ordinal);

        Assert.True(topBand >= 0, "Expected a top band anchored to the page top.");
        Assert.True(bottomBand > topBand, "Expected the bottom band after the top band in insertion order.");
        Assert.True(leftBand > bottomBand, "Expected the left band after the bottom band in insertion order.");
        Assert.True(rightBand > leftBand, "Expected the right band after the left band in insertion order.");
        Assert.True(firstTextBlock > rightBand, "Expected body text to be emitted after all background bands.");
        Assert.Contains("/ca 0.42", raw, StringComparison.Ordinal);
        Assert.Contains("/CA 0.7", raw, StringComparison.Ordinal);
        Assert.Contains("/Shading << /SH", raw, StringComparison.Ordinal);
        Assert.Contains("/SH1 sh", stream, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundBands_CanBePageScopedWithCurrentPageSize() {
        byte[] bytes = PdfDocument.Create()
            .Page(page => {
                page.Size(300, 400);
                page.Margin(30, 40, 30, 40);
                page.BackgroundTopBand(50, PdfColor.FromRgb(238, 242, 255), insetX: 12, offsetY: 8);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Scoped band"))));
            })
            .ToBytes();

        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));

        Assert.Contains("12 342 276 50 re", stream, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundShape_ValidatesAndClonesOptions() {
        var shape = PdfPageBackgroundShape.Rectangle(12, 24, 120, 48, PdfColor.FromRgb(224, 242, 254), fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.LightBlue, OfficeColor.WhiteSmoke));
        var options = new PdfOptions {
            PageBackgroundShapes = new[] { shape }
        };

        shape.X = 300;
        PdfPageBackgroundShape snapshot = Assert.Single(options.PageBackgroundShapes!);
        snapshot.X = 500;
        OfficeShape snapshotShape = snapshot.Shape;
        snapshotShape.FillColor = OfficeColor.Red;
        snapshot.Shape = snapshotShape;

        PdfPageBackgroundShape stored = Assert.Single(options.PageBackgroundShapes!);
        Assert.Equal(12, stored.X);
        Assert.NotEqual(OfficeColor.Red, stored.Shape.FillColor);
        Assert.NotNull(stored.Shape.FillGradient);
        Assert.Throws<ArgumentNullException>(() => new PdfPageBackgroundShape(null!, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBackgroundShape(OfficeShape.Rectangle(10, 10), double.NaN, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(0, 0, 10, 10, stroke: PdfColor.Black, strokeWidth: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(0, 0, 10, 10, fill: PdfColor.Black, fillOpacity: 1.1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(0, 0, 10, 10, stroke: PdfColor.Black, strokeWidth: 1, strokeOpacity: double.NaN));
        Assert.Throws<ArgumentException>(() => PdfPageBackgroundShape.TopBand(300, 400, 50, insetX: 160));
        Assert.Throws<ArgumentException>(() => PdfPageBackgroundShape.RightBand(300, 400, 80, offsetX: 240));
    }


}
