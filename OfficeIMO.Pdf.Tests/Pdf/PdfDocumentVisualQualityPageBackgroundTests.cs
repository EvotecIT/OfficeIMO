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
            .PageBorder(global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175), width: 2, inset: 30, opacity: 0.4, dashStyle: OfficeStrokeDashStyle.Dash)
            .H1("Page border proof")
            .Paragraph(p => p.Text("Body content stays inside the reusable page frame."))
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
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
        var options = new PdfCreateOptions {
            PageBorder = new PdfPageBorder {
                Color = global::OfficeIMO.OfficeColor.FromRgb(1, 2, 3),
                Width = global::OfficeIMO.OfficeLength.Points(2),
                Inset = global::OfficeIMO.OfficeLength.Points(24),
                Opacity = 0.75,
                DashStyle = OfficeStrokeDashStyle.Dot
            }
        };

        PdfPageBorder snapshot = options.PageBorder!;
        snapshot.Width = global::OfficeIMO.OfficeLength.Points(6);

        Assert.Equal(2, options.PageBorder!.Width.ToPoints());
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Width = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Inset = global::OfficeIMO.OfficeLength.Points(-1) });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageBorder { Opacity = double.NaN });

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfCreateOptions {
                    PageBorder = new PdfPageBorder {
                        Inset = global::OfficeIMO.OfficeLength.Points(400)
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

        string raw = PdfTestContentText.ReadDecoded(bytes);
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

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
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
        var options = new PdfCreateOptions {
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
        shape.FillColor = global::OfficeIMO.OfficeColor.FromRgb(234, 244, 255);
        shape.StrokeColor = global::OfficeIMO.OfficeColor.FromRgb(96, 165, 250);
        shape.StrokeWidth = 1.25;
        shape.FillOpacity = 0.34;
        shape.StrokeOpacity = 0.6;

        byte[] bytes = PdfDocument.Create()
            .BackgroundShape(new PdfPageBackgroundShape(shape, global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(640)))
            .H1("Background shape proof")
            .Paragraph(p => p.Text("Body content stays above reusable vector page decoration."))
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
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
    public void BackgroundShape_EmptyClipSuppressesPaintShadowAndTransform() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = OfficeColor.Magenta;
        shape.Shadow = new OfficeShadow(OfficeColor.Lime, 0.5, 3, 4);
        shape.Transform = OfficeTransform.Translate(10, 5);
        shape.ClipPath = OfficeClipPath.Empty();

        byte[] bytes = PdfDocument.Create()
            .BackgroundShape(new PdfPageBackgroundShape(shape, global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(640)))
            .Paragraph(p => p.Text("Empty clip background proof"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain("1 0 1 rg", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("0 1 0 rg", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("Empty clip background proof", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundShape_CanBeScopedAndClearedPerComposedPage() {
        byte[] bytes = PdfDocument.Create()
            .BackgroundRectangle(36, 640, 540, 86, global::OfficeIMO.OfficeColor.FromRgb(238, 242, 255))
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
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(300),
                PageHeight = global::OfficeIMO.OfficeLength.Points(400),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(40),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(40)
            })
            .BackgroundTopBand(50, global::OfficeIMO.OfficeColor.FromRgb(238, 242, 255), insetX: 12, offsetY: 8, stroke: global::OfficeIMO.OfficeColor.FromRgb(96, 165, 250), strokeWidth: 0.8, fillOpacity: 0.42, strokeOpacity: 0.7, fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.LightBlue, OfficeColor.WhiteSmoke))
            .BackgroundBottomBand(30, global::OfficeIMO.OfficeColor.FromRgb(240, 253, 244), insetX: 20, offsetY: 10)
            .BackgroundLeftBand(18, global::OfficeIMO.OfficeColor.FromRgb(254, 249, 195), insetY: 24, offsetX: 6)
            .BackgroundRightBand(22, global::OfficeIMO.OfficeColor.FromRgb(255, 237, 213), insetY: 30, offsetX: 9)
            .Paragraph(p => p.Text("Anchored bands"))
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
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
                page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400));
                page.Margin(global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(40));
                page.BackgroundTopBand(global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeColor.FromRgb(238, 242, 255), insetX: global::OfficeIMO.OfficeLength.Points(12), offsetY: global::OfficeIMO.OfficeLength.Points(8));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Scoped band"))));
            })
            .ToBytes();

        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));

        Assert.Contains("12 342 276 50 re", stream, StringComparison.Ordinal);
    }

    [Fact]
    public void BackgroundShape_ValidatesAndClonesOptions() {
        var shape = PdfPageBackgroundShape.Rectangle(global::OfficeIMO.OfficeLength.Points(12), global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(48), global::OfficeIMO.OfficeColor.FromRgb(224, 242, 254), fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.LightBlue, OfficeColor.WhiteSmoke));
        var options = new PdfCreateOptions {
            PageBackgroundShapes = new[] { shape }
        };

        shape.X = global::OfficeIMO.OfficeLength.Points(300);
        PdfPageBackgroundShape snapshot = Assert.Single(options.PageBackgroundShapes!);
        snapshot.X = global::OfficeIMO.OfficeLength.Points(500);
        OfficeShape snapshotShape = snapshot.Shape;
        snapshotShape.FillColor = OfficeColor.Red;
        snapshot.Shape = snapshotShape;

        PdfPageBackgroundShape stored = Assert.Single(options.PageBackgroundShapes!);
        Assert.Equal(12, stored.X.ToPoints());
        Assert.NotEqual(OfficeColor.Red, stored.Shape.FillColor);
        Assert.NotNull(stored.Shape.FillGradient);
        Assert.Throws<ArgumentNullException>(() => new PdfPageBackgroundShape(null!, global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0)));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(10), global::OfficeIMO.OfficeLength.Points(10), stroke: global::OfficeIMO.OfficeColor.Black, strokeWidth: global::OfficeIMO.OfficeLength.Points(0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(10), global::OfficeIMO.OfficeLength.Points(10), fill: global::OfficeIMO.OfficeColor.Black, fillOpacity: 1.1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageBackgroundShape.Rectangle(global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(10), global::OfficeIMO.OfficeLength.Points(10), stroke: global::OfficeIMO.OfficeColor.Black, strokeWidth: global::OfficeIMO.OfficeLength.Points(1), strokeOpacity: double.NaN));
        Assert.Throws<ArgumentException>(() => PdfPageBackgroundShape.TopBand(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(50), insetX: global::OfficeIMO.OfficeLength.Points(160)));
        Assert.Throws<ArgumentException>(() => PdfPageBackgroundShape.RightBand(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(80), offsetX: global::OfficeIMO.OfficeLength.Points(240)));
    }


}
