using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnFrameRegressionTests {
    [Fact]
    public void Columns_KeptPanelAndCaptionStayTogetherAfterWidthDrivenReadmission() {
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(612), PageHeight = global::OfficeIMO.OfficeLength.Points(792),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(72), MarginRight = global::OfficeIMO.OfficeLength.Points(72), MarginTop = global::OfficeIMO.OfficeLength.Points(72), MarginBottom = global::OfficeIMO.OfficeLength.Points(72), MaxGeneratedPages = 3 };
        var columns = new PdfMultiColumnOptions { BalanceLastPage = false, ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(100), 36), new PdfFlowColumn(PdfColumnWidth.Fixed(332), 0)
        } };
        var imageStyle = new PdfImageStyle { ScaleDownToFit = true, SpacingBefore = 0, SpacingAfter = 0 };
        byte[] image = PdfPngTestImages.CreateRgbPng(0, 160, 96);
        var line = new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false };
        var document = PdfDocument.Create(options).Columns(content => {
            content.Paragraph(p => p.Text("Lead"), style: line).Spacer(378);
            content.Panel(panel => panel.Image(image, 332, 400, style: imageStyle)
                .Image(image, 332, 400, style: imageStyle),
                new PdfPanelStyle { PaddingX = 0, PaddingY = 0, BorderWidth = 0,
                    SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true, KeepWithNext = true });
            content.Paragraph(p => p.Text("Caption"), style: line);
        }, columns);
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(new[] { "Lead", "Caption" }, pdf.GetPages()
            .SelectMany(page => page.GetWords()).Select(word => word.Text));
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Lead");
        var caption = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "Caption");

        // The wider destination cannot contain the two authored images as a
        // kept group. A later narrow frame can: neither image may be dropped,
        // split from its peer, clipped or separated from the following caption.
        var images = PdfImageExtractor.ExtractImagePlacements(bytes)
            .OrderByDescending(placement => placement.Y).ToArray();
        Assert.Equal(2, images.Length);
        double imageWidth = 100D;
        double imageHeight = 400D * imageWidth / 332D;
        Assert.All(images, placement => {
            Assert.Equal(2, placement.PageNumber);
            Assert.Equal(options.MarginLeft.ToPoints(), placement.X, 3);
            Assert.Equal(imageWidth, placement.Width, 3);
            Assert.Equal(imageHeight, placement.Height, 3);
            Assert.InRange(placement.Y, options.MarginBottom.ToPoints(),
                options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - placement.Height + .001D);
            Assert.True(placement.X + placement.Width <= options.PageWidth.ToPoints() - options.MarginRight.ToPoints() + .001D);
            if (placement.Clip is { } clip) {
                Assert.True(clip.IsRectangle && clip.IsExact);
                double top = options.PageHeight.ToPoints() - placement.Y - placement.Height;
                double visibleWidth = Math.Max(0D, Math.Min(placement.X + placement.Width, clip.X + clip.Width) -
                    Math.Max(placement.X, clip.X));
                double visibleHeight = Math.Max(0D, Math.Min(top + placement.Height, clip.Y + clip.Height) -
                    Math.Max(top, clip.Y));
                Assert.Equal(placement.Width, visibleWidth, 3);
                Assert.Equal(placement.Height, visibleHeight, 3);
            }
        });
        Assert.Equal(options.PageHeight.ToPoints() - options.MarginTop.ToPoints(), images[0].Y + images[0].Height, 3);
        Assert.Equal(images[0].Y, images[1].Y + images[1].Height, 3);
        Assert.Empty(pdf.GetPage(1).GetImages());
        Assert.Equal(2, pdf.GetPage(2).GetImages().Count());
        Assert.InRange(caption.BoundingBox.Left, options.MarginLeft.ToPoints() - .001D, options.MarginLeft.ToPoints() + imageWidth);
        Assert.InRange(caption.BoundingBox.Top, images[1].Y - 20D, images[1].Y + .001D);
        Assert.True(caption.BoundingBox.Bottom >= options.MarginBottom.ToPoints());
    }

    [Fact]
    public void Columns_LargeSpacerTraversesEveryColumnWithoutCountingColumnsAsPages() {
        var options = Options(); options.MaxGeneratedPages = 2;
        byte[] bytes = PdfDocument.Create(options).Columns(content => {
            content.Spacer(650);
            content.Paragraph(p => p.Text("AfterSpacer"), style: Paragraph());
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var word = Assert.Single(pdf.GetPage(pdf.NumberOfPages).GetWords());
        Assert.Equal("AfterSpacer", word.Text);
        Assert.InRange(word.BoundingBox.Left, 19.9, 20.1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Columns_ImageUsesFullPhysicalContinuationCapacityAfterAPrelude(bool scaleDown) {
        byte[] bytes = PdfDocument.Create(Options()).Paragraph(p => p.Text(Lines("Prelude", 9)), style: Paragraph())
            .Columns(content => content.Image(PdfPngTestImages.CreateRgbPng(255, 0, 0), 100, 120,
                style: new PdfImageStyle { ScaleDownToFit = scaleDown, SpacingBefore = 0, SpacingAfter = 0 }),
                new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        var image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(2, image.PageNumber);
        Assert.InRange(image.Height, 119.9, 120.1);
    }

    [Fact]
    public void Columns_FixedFieldSkipsUnusedPartialColumnsUntilItFits() {
        byte[] bytes = PdfDocument.Create(Options()).Paragraph(p => p.Text(Lines("Prelude", 9)), style: Paragraph())
            .Columns(content => content.TextField("TallField", width: global::OfficeIMO.OfficeLength.Points(100), height: global::OfficeIMO.OfficeLength.Points(120), spacingAfter: global::OfficeIMO.OfficeLength.Points(0)),
                new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        var widget = Assert.Single(PdfInspector.Inspect(bytes).FormFieldsByName["TallField"].Widgets);
        Assert.Equal(2, widget.PageNumber);
        Assert.InRange(widget.Y1, 159.9, 160.1);
    }

    [Fact]
    public void Columns_ScaledImageRecomputesItsBoxInTheNarrowerContinuation() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Spacer(240);
            content.Image(PdfPngTestImages.CreateRgbPng(255, 0, 0), 200, 80,
                style: new PdfImageStyle { ScaleDownToFit = true, SpacingBefore = 0, SpacingAfter = 0 });
        }, Unequal()).ToBytes();
        var image = Assert.Single(PdfDocument.Load(bytes).Images.Placements());
        Assert.Equal(1, image.PageNumber);
        Assert.InRange(image.Width, 99.9, 100.1);
        Assert.InRange(image.Height, 39.9, 40.1);
    }

    [Fact]
    public void Columns_FixedFieldSkipsANarrowContinuationRatherThanOverflowingIt() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Paragraph(p => p.Text("Prelude"), style: Paragraph());
            content.Spacer(220);
            content.TextField("WideField", width: global::OfficeIMO.OfficeLength.Points(200), height: global::OfficeIMO.OfficeLength.Points(80), spacingAfter: global::OfficeIMO.OfficeLength.Points(0));
        }, Unequal()).ToBytes();
        var widget = Assert.Single(PdfInspector.Inspect(bytes).FormFieldsByName["WideField"].Widgets);
        Assert.Equal(2, widget.PageNumber);
        Assert.InRange(widget.X1, 19.9, 20.1);
        Assert.InRange(widget.X2, 219.9, 220.1);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("panel")]
    [InlineData("keep")]
    public void Columns_BalancingSuppressesParagraphSpacingBeforeAtAColumnStart(string kind) {
        var style = Paragraph(); style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(200);
        if (kind == "keep") style.KeepTogether = true;
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            if (kind == "panel") content.Panel(panel => panel.Paragraph(p => p.Text(Lines("Line", 8)), style: style),
                new PdfPanelStyle { PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0 });
            else content.Paragraph(p => p.Text(Lines("Line", 8)), style: style);
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceKeptParagraphLines = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.InRange(words.Single(word => word.Text == "Line04").BoundingBox.Left, 19.9, 20.1);
        Assert.InRange(words.Single(word => word.Text == "Line05").BoundingBox.Left, 239.9, 240.1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Columns_ContinuedNoWrapTableTextShrinksToTheNewFrame(bool explicitSize) {
        string text = string.Join("\n", Enumerable.Range(1, 20).Select(i => $"LongMarker{i:D2}UnwrappedSuffix"));
        var cell = (explicitSize ? new PdfTableCell(new[] { new PdfTextRun(text, fontSize: 12) }) : new PdfTableCell(text)).WithNoWrap();
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Table(new[] { new[] { cell } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
                FontSize = global::OfficeIMO.OfficeLength.Points(12), LineHeight = 20D / 12D, ShrinkTextToFit = true, MinimumShrinkFontSize = global::OfficeIMO.OfficeLength.Points(4),
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }), Unequal()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var continued = pdf.GetPage(1).Letters.Where(letter => letter.StartBaseLine.X >= 339.9).ToArray();
        Assert.NotEmpty(continued);
        Assert.True(continued.Max(letter => letter.FontSize) < 12);
        Assert.All(continued, letter => Assert.InRange(letter.BoundingBox.Right, 340, 440.1));
        string rendered = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        foreach (int index in Enumerable.Range(1, 20)) Assert.Contains($"LongMarker{index:D2}UnwrappedSuffix", rendered);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Columns_ContinuedCellTextFitsTheNarrowFrameWithoutRowShrinking(bool explicitSize, bool paragraphs) {
        string text = string.Join("\n", Enumerable.Range(1, 20).Select(i => $"LongMarker{i:D2}UnwrappedSuffix"));
        var runs = new[] { new PdfTextRun(text, fontSize: explicitSize ? 12 : null) };
        var cell = (paragraphs
            ? new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs) })
            : new PdfTableCell(runs)).WithNoWrap().WithShrinkTextToFit();
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Table(new[] { new[] { cell } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
                FontSize = global::OfficeIMO.OfficeLength.Points(12), LineHeight = 20D / 12D, ShrinkTextToFit = false, MinimumShrinkFontSize = global::OfficeIMO.OfficeLength.Points(4),
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }), Unequal()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var continued = pdf.GetPage(1).Letters.Where(letter => letter.StartBaseLine.X >= 339.9).ToArray();
        Assert.NotEmpty(continued);
        Assert.All(continued, letter => Assert.True(letter.FontSize < 12));
        Assert.All(continued, letter => Assert.InRange(letter.BoundingBox.Right, 340, 440.1));
        string rendered = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        foreach (int index in Enumerable.Range(1, 20)) Assert.Contains($"LongMarker{index:D2}UnwrappedSuffix", rendered);
    }

    [Fact]
    public void Columns_FixedObjectCannotRetryPastAContainerMaximumWidth() {
        var options = Options(); options.MaxGeneratedPages = 2;
        var document = PdfDocument.Create(options).Columns(content => content.Panel(panel =>
            panel.TextField("Impossible", width: global::OfficeIMO.OfficeLength.Points(80), height: global::OfficeIMO.OfficeLength.Points(20)),
            new PdfPanelStyle { MaxWidth = 60, PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0 }), Unequal());
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData("image")]
    [InlineData("shape")]
    [InlineData("drawing")]
    public void Columns_WidthDrivenObjectPlacementRemeasuresItsKeepInTheDestination(string kind) {
        var kept = Paragraph(); kept.KeepTogether = true;
        var shape = OfficeShape.Rectangle(200, 40); shape.FillColor = OfficeColor.Red;
        var document = PdfDocument.Create(Options());
        document.Content.Paragraph(p => p.Text(Lines("Prelude", 9)), style: Paragraph());
        document.Content.Columns(content => {
            if (kind == "image") content.Image(PdfPngTestImages.CreateRgbPng(255, 0, 0), 200, 40,
                style: new PdfImageStyle { ScaleDownToFit = false, KeepWithNext = true, SpacingBefore = 0, SpacingAfter = 0 });
            else if (kind == "shape") content.Shape(shape,
                style: new PdfDrawingStyle { KeepWithNext = true, SpacingBefore = 0, SpacingAfter = 0 });
            else content.Drawing(new OfficeDrawing(200, 40).AddShape(shape, 0, 0),
                style: new PdfDrawingStyle { KeepWithNext = true, SpacingBefore = 0, SpacingAfter = 0 });
            content.Paragraph(p => p.Text(string.Join("\n", Enumerable.Repeat("Following words for kept text", 6))), style: kept);
        }, new PdfMultiColumnOptions { BalanceLastPage = false, ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(300))
        }});
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Following", pdf.GetPage(1).Text);
        Assert.Contains("Following", pdf.GetPage(2).Text);
        if (kind == "image") Assert.Equal(2, Assert.Single(PdfDocument.Load(bytes).Images.Placements()).PageNumber);
        else {
            Assert.Empty(pdf.GetPage(1).Paths);
            Assert.NotEmpty(pdf.GetPage(2).Paths);
        }
    }

    [Fact]
    public void Columns_ParagraphTableKeepChainRetainsConditionalBeforeSpacingDuringBalancing() {
        var before = Paragraph(); before.SpacingBefore = global::OfficeIMO.OfficeLength.Points(200); before.KeepWithNext = true;
        var following = Paragraph(); following.KeepTogether = true;
        var document = PdfDocument.Create(Options());
        document.Content.Columns(content => {
            content.Paragraph(p => p.Text("Preceding"), style: Paragraph());
            content.Paragraph(p => p.Text("KeptLeader"), style: before);
            content.Table(new[] { new[] { new PdfTableCell("TableLine1\nTableLine2") } }, style: new PdfTableStyle {
                KeepWithNext = true, HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), FontSize = global::OfficeIMO.OfficeLength.Points(12),
                LineHeight = 20D / 12D, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), BorderWidth = global::OfficeIMO.OfficeLength.Points(0),
                MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0
            });
            content.Paragraph(p => p.Text("Following"), style: Paragraph());
            content.Paragraph(p => p.Text(Lines("Trailing", 5)), style: following);
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceTableRowLines = true, BalanceParagraphLines = false });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.InRange(words.Single(word => word.Text == "KeptLeader").BoundingBox.Left, 239.9, 240.1);
        Assert.Contains(words, word => word.Text == "Trailing05");
    }

    private static PdfCreateOptions Options() => new() {
        BackgroundColor = null,
        PageWidth = global::OfficeIMO.OfficeLength.Points(460), PageHeight = global::OfficeIMO.OfficeLength.Points(300), MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20),
        MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12), CompressContentStreams = false
    };
    private static PdfParagraphStyle Paragraph() => new() {
        LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false
    };
    private static PdfMultiColumnOptions Unequal() => new() {
        BalanceLastPage = false, ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(100))
        }
    };
    private static string Lines(string prefix, int count) => string.Join("\n", Enumerable.Range(1, count).Select(i => $"{prefix}{i:D2}"));
}
