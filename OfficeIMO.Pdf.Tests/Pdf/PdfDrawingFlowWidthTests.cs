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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TallDrawingCanFitTheNarrowerSequentialColumn(bool keepWithNext) {
        var document = PdfDocument.Create(Options(460, 300));
        PdfDrawingStyle style = Constrained();
        style.KeepWithNext = keepWithNext;
        document.Content.Columns(content => {
            content.Drawing(ImageDrawing(400, 800), style: style, linkUri: ObjectLink);
            content.Paragraph(p => p.Text("Caption"), style: Paragraph());
        }, UnequalColumns());
        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(340, image.X, 2);
        Assert.Equal(100, image.Width, 2);
        Assert.Equal(200, image.Height, 2);
        Assert.Equal(100, Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink)).Width, 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PanelDrawingDoesNotRequireAnUnusedNarrowColumnToHaveValidPadding(bool keepWithNext) {
        var document = PdfDocument.Create(Options(460, 300));
        PdfDrawingStyle style = Constrained();
        style.KeepWithNext = keepWithNext;
        document.Content.Columns(content => content.Panel(panel => {
            panel.Drawing(ImageDrawing(100, 40), style: style, linkUri: ObjectLink);
            panel.Paragraph(p => p.Text("Caption"), style: Paragraph());
        }, PlainPanel(60)), UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(80, image.X, 2);
        Assert.Equal(100, image.Width, 2);
        Assert.Equal(40, image.Height, 2);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.X, link.X1, 2);
        Assert.Equal(image.Width, link.Width, 2);
        Assert.Equal(image.Height, link.Height, 2);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void TallPanelDrawingCanFitTheNarrowerSequentialColumn(bool keepTogether, bool nested, bool keepWithNext) {
        var document = PdfDocument.Create(Options(460, 300));
        PdfDrawingStyle drawingStyle = Constrained();
        drawingStyle.KeepWithNext = keepWithNext;
        PdfPanelStyle panelStyle = PlainPanel(nested ? 5 : 10);
        panelStyle.KeepTogether = keepTogether;
        document.Content.Columns(content => content.Panel(panel => {
            Action<PdfContentBuilder> addDrawing = target => {
                target.Drawing(ImageDrawing(400, 800), style: drawingStyle, linkUri: ObjectLink);
                target.Paragraph(p => p.Text("Caption"), style: Paragraph());
            };
            if (nested) panel.Panel(addDrawing, PlainPanel(5));
            else addDrawing(panel);
        }, panelStyle), UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(350, image.X, 2);
        Assert.Equal(80, image.Width, 2);
        Assert.Equal(160, image.Height, 2);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.X, link.X1, 2);
        Assert.Equal(image.Width, link.Width, 2);
        Assert.Equal(image.Height, link.Height, 2);
        using var pdf = PdfPigDocument.Open(bytes);
        var caption = Assert.Single(pdf.GetPage(image.PageNumber).GetWords(), word => word.Text == "Caption");
        Assert.Equal(image.X, caption.BoundingBox.Left, 2);
        Assert.True(caption.BoundingBox.Top <= image.Y + .01D);
    }

    [Theory]
    [InlineData(2000, 0)]
    [InlineData(800, 300)]
    public void PanelDrawingThatCannotFitAnyColumnFailsWithoutRetryingPages(double height, double spacingAfter) {
        PdfOptions options = Options(460, 300);
        options.MaxGeneratedPages = 2;
        var document = PdfDocument.Create(options);
        PdfDrawingStyle style = Constrained();
        style.SpacingAfter = spacingAfter;
        document.Content.Columns(content => content.Panel(panel =>
            panel.Drawing(ImageDrawing(400, height), style: style), PlainPanel(10)), UnequalColumns());

        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(2, 240, false)]
    [InlineData(6, 80, true)]
    public void KeptPanelMeasuresItsWholeDrawingSequenceAtTheNarrowerColumnWidth(int drawingCount, double drawingHeight, bool partialFirstPage) {
        var options = Options(460, 300);
        options.MaxGeneratedPages = 2;
        var document = PdfDocument.Create(options);
        if (partialFirstPage) {
            PdfParagraphStyle prelude = Paragraph();
            prelude.FontSize = 20;
            document.Content.Paragraph(p => p.Text("Prelude"), style: prelude).Spacer(180);
        }
        PdfPanelStyle panelStyle = PlainPanel(10);
        panelStyle.KeepTogether = true;
        OfficeDrawing drawing = ImageDrawing(400, drawingHeight);
        document.Content.Columns(content => content.Panel(panel => {
            for (int index = 0; index < drawingCount; index++)
                panel.Drawing(drawing, style: Constrained(), linkUri: ObjectLink);
            panel.Paragraph(p => p.Text("Caption"), style: Paragraph());
        }, panelStyle), UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement[] images = PdfDocument.Load(bytes).Images.Placements().ToArray();
        Assert.Equal(drawingCount, images.Length);
        Assert.All(images, image => {
            Assert.Equal(partialFirstPage ? 2 : 1, image.PageNumber);
            Assert.Equal(350, image.X, 2);
            Assert.Equal(80, image.Width, 2);
            Assert.Equal(drawingHeight / 5D, image.Height, 2);
        });
        PdfLinkAnnotation[] links = PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink).ToArray();
        Assert.Equal(drawingCount, links.Length);
        Assert.All(links, link => {
            Assert.Equal(partialFirstPage ? 2 : 1, link.PageNumber);
            Assert.Equal(80, link.Width, 2);
            Assert.Equal(drawingHeight / 5D, link.Height, 2);
        });
        using var pdf = PdfPigDocument.Open(bytes);
        if (partialFirstPage) {
            Assert.Contains("Prelude", pdf.GetPage(1).Text);
            Assert.DoesNotContain("Caption", pdf.GetPage(1).Text);
        }
        var caption = Assert.Single(pdf.GetPage(images[0].PageNumber).GetWords(), word => word.Text == "Caption");
        Assert.Equal(images[0].X, caption.BoundingBox.Left, 2);
        Assert.True(caption.BoundingBox.Top <= images.Min(image => image.Y) + .01D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KeptStaticFlowMeasuresItsDrawingAtTheNarrowerColumnWidth(bool partialFirstPage) {
        var options = Options(460, 300);
        options.MaxGeneratedPages = 2;
        var document = PdfDocument.Create(options);
        if (partialFirstPage) {
            PdfParagraphStyle prelude = Paragraph();
            prelude.FontSize = 20;
            document.Content.Paragraph(p => p.Text("Prelude"), style: prelude).Spacer(180);
        }
        document.Content.Columns(content => content.Flow(flow => {
            flow.Drawing(ImageDrawing(400, 800), style: Constrained(), linkUri: ObjectLink);
            flow.Paragraph(p => p.Text("Caption"), style: Paragraph());
        }, new PdfFlowOptions { KeepTogether = true }), UnequalColumns());

        byte[] bytes = document.ToBytes();
        PdfImagePlacement image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(partialFirstPage ? 2 : 1, image.PageNumber);
        Assert.Equal(340, image.X, 2);
        Assert.Equal(100, image.Width, 2);
        Assert.Equal(200, image.Height, 2);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(ObjectLink));
        Assert.Equal(image.PageNumber, link.PageNumber);
        Assert.Equal(image.X, link.X1, 2);
        Assert.Equal(image.Y, link.Y1, 2);
        Assert.Equal(image.Width, link.Width, 2);
        Assert.Equal(image.Height, link.Height, 2);
        using var pdf = PdfPigDocument.Open(bytes);
        if (partialFirstPage) {
            Assert.Contains("Prelude", pdf.GetPage(1).Text);
            Assert.DoesNotContain("Caption", pdf.GetPage(1).Text);
        }
        var caption = Assert.Single(pdf.GetPage(image.PageNumber).GetWords(), word => word.Text == "Caption");
        Assert.Equal(image.X, caption.BoundingBox.Left, 2);
        Assert.True(caption.BoundingBox.Top <= image.Y + .01D);
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

    private static PdfPanelStyle PlainPanel(double paddingX) => new() {
        PaddingX = paddingX, PaddingY = 0, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0
    };

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
