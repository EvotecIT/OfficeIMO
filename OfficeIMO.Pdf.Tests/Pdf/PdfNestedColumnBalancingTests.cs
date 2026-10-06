using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNestedColumnBalancingTests {
    [Theory]
    [InlineData("semantic", "paragraph", 21)]
    [InlineData("semantic", "paragraph", 43)]
    [InlineData("semantic", "list", 21)]
    [InlineData("semantic", "list", 43)]
    [InlineData("semantic", "table", 21)]
    [InlineData("semantic", "table", 43)]
    [InlineData("flow", "paragraph", 21)]
    [InlineData("flow", "paragraph", 43)]
    [InlineData("flow", "list", 21)]
    [InlineData("flow", "list", 43)]
    [InlineData("flow", "table", 21)]
    [InlineData("flow", "table", 43)]
    [InlineData("panel", "paragraph", 21)]
    [InlineData("panel", "paragraph", 43)]
    [InlineData("panel", "list", 21)]
    [InlineData("panel", "list", 43)]
    [InlineData("panel", "table", 21)]
    [InlineData("panel", "table", 43)]
    public void Columns_BalancesNestedContentAndItsFinalPage(string wrapper, string kind, int count) {
        var document = PdfDocument.Create(Options()).Columns(columns => {
            Action<PdfContentBuilder> compose = content => AddContent(content, kind, count);
            if (wrapper == "semantic") columns.Semantic(PdfSemanticRole.Section, compose);
            else if (wrapper == "flow") columns.Flow(compose);
            else columns.Panel(compose, PanelStyle());
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceTableRowLines = true })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        int finalPage = count == 21 ? 1 : 2;
        int firstFinal = count == 21 ? 1 : wrapper == "panel" ? 31 : 33;
        int leftLast = count == 21 ? 11 : wrapper == "panel" ? 37 : 38;
        Assert.Equal(finalPage, pdf.NumberOfPages);
        Assert.InRange(X(pdf, finalPage, $"Marker{leftLast:D3}"), 40, kind == "list" ? 80 : 40.1);
        Assert.InRange(X(pdf, finalPage, $"Marker{leftLast + 1:D3}"), 260, kind == "list" ? 300 : 260.1);
        // Bottom padding at a full column is bounded by its remaining height, as in ordinary container flow.
        double expectedAfter = (leftLast - firstFinal + 1) * 20;
        Assert.InRange(Y(pdf, finalPage, $"Marker{firstFinal:D3}") - Y(pdf, finalPage, "AfterColumns"), expectedAfter - .1, expectedAfter + .1);
        AssertMarkers(pdf, count);
    }

    [Fact]
    public void Columns_NestedContinuationIncludesFollowingSiblingsAtEveryAncestor() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => {
            columns.Semantic(PdfSemanticRole.Section, section => {
                section.Flow(flow => {
                    AddContent(flow, "paragraph", 37);
                    flow.Paragraph(p => p.Text("FlowSibling"), style: ParagraphStyle());
                });
                section.Paragraph(p => p.Text("SectionSibling"), style: ParagraphStyle());
            });
            columns.Paragraph(p => p.Text("RootSibling"), style: ParagraphStyle());
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 2, "Marker036"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, "Marker037"), 259.9, 260.1);
        foreach (string sibling in new[] { "FlowSibling", "SectionSibling", "RootSibling" }) {
            Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == sibling);
            Assert.InRange(X(pdf, 2, sibling), 259.9, 260.1);
        }
        AssertMarkers(pdf, 37);
    }

    [Fact]
    public void Columns_NestedPanelsRepeatTheAccumulatedPaddingOnTheBalancedContinuation() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => columns.Panel(outer => {
            var inner = PanelStyle(); inner.PaddingY = 5;
            outer.Panel(content => content.Flow(flow => AddContent(flow, "paragraph", 43)), inner);
        }, PanelStyle()), new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 2, "Marker037"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, "Marker038"), 259.9, 260.1);
        Assert.InRange(Math.Abs(Y(pdf, 2, "Marker031") - Y(pdf, 2, "Marker038")), 0, .1);
        AssertMarkers(pdf, 43);
    }

    [Fact]
    public void Columns_NestedPanelMeasurementUsesItsMaximumWidthAndHorizontalPadding() {
        var style = PanelStyle(); style.MaxWidth = 100; style.PaddingX = 10; style.Align = PdfAlign.Center;
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => columns.Panel(panel => panel.Paragraph(p =>
            p.Text(string.Join(" ", Enumerable.Range(1, 21).Select(index => $"Marker{index:D3}"))), style: ParagraphStyle()), style),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Marker011"), 99.9, 100.1);
        Assert.InRange(X(pdf, 1, "Marker012"), 319.9, 320.1);
        AssertMarkers(pdf, 21);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("panel")]
    public void Columns_KeepsAnEntireConstrainedNestedGroupInOneColumn(string wrapper) {
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => {
            columns.Paragraph(p => p.Text("Prelude1\nPrelude2\nPrelude3"), style: ParagraphStyle());
            if (wrapper == "flow") columns.Flow(content => AddContent(content, "paragraph", 5), new PdfFlowOptions { KeepTogether = true });
            else {
                var style = PanelStyle(); style.KeepTogether = true;
                columns.Panel(content => AddContent(content, "paragraph", 5), style);
            }
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Prelude3"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "Marker001"), 259.9, 260.1);
        Assert.InRange(X(pdf, 1, "Marker005"), 259.9, 260.1);
        AssertMarkers(pdf, 5);
    }

    [Fact]
    public void Columns_BalancesRowFragmentsInsideAPaddedSemanticContainer() {
        var style = new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12,
            LineHeight = 20D / 12D, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0,
            MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0 };
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => columns.Semantic(PdfSemanticRole.Section,
            section => section.Panel(panel => panel.Table(new[] { new[] { new PdfTableCell(string.Join("\n",
                Enumerable.Range(1, 43).Select(index => $"Marker{index:D3}"))) } }, style: style), PanelStyle())),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceTableRowLines = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 2, "Marker037"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, "Marker038"), 259.9, 260.1);
        AssertMarkers(pdf, 43);
    }

    [Fact]
    public void Columns_NestedRowKeepRulesUseTheContentHeightAfterContainerPadding() {
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, 16).Select(index => $"Marker{index:D3}"))) };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false, keepTogether: true) });
        var panel = PanelStyle(); panel.PaddingY = 20;
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => columns.Panel(content => content.Table(new[] { new[] { cell } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 20D / 12D,
                BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0, MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0 }), panel),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceTableRowLines = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Marker008"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "Marker009"), 259.9, 260.1);
        AssertMarkers(pdf, 16);
    }

    private static void AddContent(PdfContentBuilder content, string kind, int count) {
        string[] markers = Enumerable.Range(1, count).Select(index => $"Marker{index:D3}").ToArray();
        if (kind == "paragraph") content.Paragraph(p => p.Text(string.Join("\n", markers)), style: ParagraphStyle());
        else if (kind == "list") content.RichNumbered(markers.Select(marker => new PdfListItem(marker)), style: new PdfListStyle {
            LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingBefore = 0, SpacingAfter = 0
        });
        else content.Table(markers.Select(marker => new[] { new PdfTableCell(marker) }), style: new PdfTableStyle {
            HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 20D / 12D,
            BorderWidth = 0, RowSeparatorWidth = 0, SpacingBefore = 0, SpacingAfter = 0,
            MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0
        });
    }

    private static void AssertMarkers(PdfPigDocument pdf, int count) {
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        Assert.Equal(Enumerable.Range(1, count), System.Text.RegularExpressions.Regex.Matches(text, @"Marker(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value)));
    }

    private static double X(PdfPigDocument pdf, int page, string marker) => Word(pdf, page, marker).BoundingBox.Left;
    private static double Y(PdfPigDocument pdf, int page, string marker) => Word(pdf, page, marker).BoundingBox.Bottom;
    private static UglyToad.PdfPig.Content.Word Word(PdfPigDocument pdf, int page, string marker) =>
        pdf.GetPage(page).GetWords().Single(word => word.Text == marker || word.Text.EndsWith("." + marker, StringComparison.Ordinal));
    private static PdfOptions Options() => new() { PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
        MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, CompressContentStreams = false };
    private static PdfParagraphStyle ParagraphStyle() => new() { LineSpacing = PdfLineSpacing.Exactly(20),
        SpacingBefore = 0, SpacingAfter = 0, WidowControl = false };
    private static PdfPanelStyle PanelStyle() => new() { PaddingX = 0, PaddingY = 10, SpacingBefore = 0, SpacingAfter = 0 };
}
