using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnFrameRegressionTests {
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
            .Columns(content => content.TextField("TallField", width: 100, height: 120, spacingAfter: 0),
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
            content.TextField("WideField", width: 200, height: 80, spacingAfter: 0);
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
        var style = Paragraph(); style.SpacingBefore = 200;
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
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
                FontSize = 12, LineHeight = 20D / 12D, ShrinkTextToFit = true, MinimumShrinkFontSize = 4,
                BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0 }), Unequal()).ToBytes();
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
            panel.TextField("Impossible", width: 80, height: 20),
            new PdfPanelStyle { MaxWidth = 60, PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0 }), Unequal());
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private static PdfOptions Options() => new() {
        PageWidth = 460, PageHeight = 300, MarginLeft = 20, MarginRight = 20,
        MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, CompressContentStreams = false
    };
    private static PdfParagraphStyle Paragraph() => new() {
        LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
    };
    private static PdfMultiColumnOptions Unequal() => new() {
        BalanceLastPage = false, ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(100))
        }
    };
    private static string Lines(string prefix, int count) => string.Join("\n", Enumerable.Range(1, count).Select(i => $"{prefix}{i:D2}"));
}
