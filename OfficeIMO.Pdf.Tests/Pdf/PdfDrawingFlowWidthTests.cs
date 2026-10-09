using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingFlowWidthTests {
    private const string ObjectLink = "https://example.test/drawing";
    private const string EmbeddedLink = "https://example.test/embedded";

    [Fact]
    public void ConstraintUsesPaddedPanelWidthAndPreservesObjectAndEmbeddedLinkBounds() {
        OfficeDrawing drawing = ImageDrawing(400, 160)
            .AddLink(EmbeddedLink, 100, 40, 200, 80);
        var document = PdfDocument.Create(Options(300, 220));
        document.Content.Panel(content => content.Drawing(drawing, style: Constrained(), linkUri: ObjectLink),
            new PdfPanelStyle { MaxWidth = 180, PaddingX = 10, PaddingY = 0, BorderWidth = 0,
                SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true });

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(30, image.X, 2);
        Assert.Equal(160, image.Width, 2);
        Assert.Equal(64, image.Height, 2);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation outer = Assert.Single(info.GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.X, outer.X1, 2);
        Assert.Equal(image.Width, outer.Width, 2);
        Assert.Equal(image.Height, outer.Height, 2);
        PdfLinkAnnotation inner = Assert.Single(info.GetLinkAnnotationsByUri(EmbeddedLink));
        Assert.Equal(70, inner.X1, 2);
        Assert.Equal(80, inner.Width, 2);
        Assert.Equal(32, inner.Height, 2);
        Assert.InRange(inner.Y1, outer.Y1, outer.Y2);
        Assert.InRange(inner.Y2, outer.Y1, outer.Y2);
    }

    [Fact]
    public void ConstraintRecomputesTheBoxAfterMovingToANarrowerSequentialColumn() {
        var document = PdfDocument.Create(Options(460, 300));
        document.Content.Columns(content => {
            content.Spacer(240);
            content.Drawing(ImageDrawing(200, 80), style: Constrained(), linkUri: ObjectLink);
        }, UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(340, image.X, 2);
        Assert.Equal(100, image.Width, 2);
        Assert.Equal(40, image.Height, 2);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.X, link.X1, 2);
        Assert.Equal(image.Width, link.Width, 2);
        Assert.Equal(image.Height, link.Height, 2);
    }

    [Fact]
    public void KeptRowUsesConstrainedDrawingHeightForMeasurementAndPlacement() {
        var document = PdfDocument.Create(Options(260, 220));
        document.Content.Spacer(80);
        document.Content.Row(row => row.Gap(20).Style(new PdfRowStyle { KeepTogether = true })
            .FixedColumn(100, content => content.Drawing(ImageDrawing(400, 320), style: Constrained(), linkUri: ObjectLink)
                .Paragraph(p => p.Text("Caption"), style: Paragraph()))
            .FixedColumn(100, content => content.Paragraph(p => p.Text("Adjacent"), style: Paragraph())));

        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Caption", pdf.GetPage(1).Text);
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(100, image.Width, 2);
        Assert.Equal(80, image.Height, 2);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.Width, link.Width, 2);
        Assert.Equal(image.Height, link.Height, 2);
    }

    [Fact]
    public void KeepWithNextRemeasuresTheDrawingAndCaptionInTheDestinationColumn() {
        var document = PdfDocument.Create(Options(460, 300));
        document.Content.Paragraph(p => p.Text("Prelude"), style: Paragraph()).Spacer(160);
        PdfDrawingStyle style = Constrained();
        style.KeepWithNext = true;
        document.Content.Columns(content => {
            content.Drawing(ImageDrawing(200, 160), style: style);
            content.Paragraph(p => p.Text("KeptCaption"), style: Paragraph());
        }, UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(2, image.PageNumber);
        Assert.Equal(20, image.X, 2);
        Assert.Equal(200, image.Width, 2);
        Assert.Equal(160, image.Height, 2);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Prelude", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeptCaption", pdf.GetPage(1).Text);
        var caption = Assert.Single(pdf.GetPage(image.PageNumber).GetWords(), word => word.Text == "KeptCaption");
        Assert.Equal(image.X, caption.BoundingBox.Left, 2);
        Assert.True(caption.BoundingBox.Top <= image.Y + .01D);
    }

    [Fact]
    public void FittingSnapshotsTheStyleAndRetainsTheReusableSourceWithoutEnlargingIt() {
        OfficeDrawing drawing = ImageDrawing(400, 160);
        PdfDrawingStyle style = Constrained();
        var narrow = PdfDocument.Create(Options(240, 220));
        narrow.Content.Drawing(drawing, style: style);
        style.ConstrainToContentWidth = false;

        PdfImagePlacement fitted = Assert.Single(PdfDocument.Load(narrow.ToBytes()).Images.Placements());
        Assert.Equal(200, fitted.Width, 2);
        Assert.Equal(80, fitted.Height, 2);
        var wide = PdfDocument.Create(Options(600, 300));
        wide.Content.Drawing(drawing, style: Constrained());
        PdfImagePlacement original = Assert.Single(PdfDocument.Load(wide.ToBytes()).Images.Placements());
        Assert.Equal(400, original.Width, 2);
        Assert.Equal(160, original.Height, 2);
        Assert.Equal(400, drawing.Width);
        Assert.Equal(160, drawing.Height);
    }

    private static OfficeDrawing ImageDrawing(double width, double height) =>
        new OfficeDrawing(width, height).AddImage(PdfPngTestImages.CreateRgbPng(255, 0, 0), "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, width, height)));

    private static PdfDrawingStyle Constrained() => new() { ConstrainToContentWidth = true };

    private static PdfOptions Options(double width, double height) => new() {
        PageWidth = width, PageHeight = height, MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20
    };

    private static PdfParagraphStyle Paragraph() => new() {
        LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
    };

    private static PdfMultiColumnOptions UnequalColumns() => new() {
        BalanceLastPage = false, ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(100))
        }
    };
}
