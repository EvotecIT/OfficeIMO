using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnStartContinuationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Columns_ExplicitKeptParagraphBalancingPolicyRetainsPhysicalPageKeeps(bool allowBalancing, bool partialPage) {
        var document = PdfDocument.Create(Options());
        if (partialPage) document.Paragraph(p => p.Text(Markers("Prelude", 12)), style: ParagraphStyle());
        var options = ColumnOptions(); options.BalanceKeptParagraphLines = allowBalancing;
        document.Columns(columns => AddKeptContent(columns, "paragraph", 13), options);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        int page = partialPage ? 2 : 1;
        Assert.Equal(page, pdf.NumberOfPages);
        if (partialPage) Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.Contains("Marker"));
        Assert.InRange(X(pdf, page, "Marker007"), 39.9, 40.1);
        Assert.InRange(X(pdf, page, "Marker008"), allowBalancing ? 259.9 : 39.9, allowBalancing ? 260.1 : 40.1);
        AssertMarkers(pdf, 13);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    public void Columns_FirstLineThatExceedsAPartialColumnCanUseTheNextPhysicalPage(string kind) {
        byte[] bytes = PdfDocument.Create(Options()).Paragraph(p => p.Text(Markers("Prelude", 12)), style: ParagraphStyle())
            .Columns(columns => {
                if (kind == "paragraph") {
                    var style = ParagraphStyle(); style.LineSpacing = PdfLineSpacing.Exactly(120);
                    columns.Paragraph(p => p.Text("Marker001"), style: style);
                } else columns.RichNumbered(new[] { new PdfListItem("Marker001") }, style: new PdfListStyle {
                    LineSpacing = PdfLineSpacing.Exactly(120), SpacingBefore = 0, SpacingAfter = 0, ItemSpacing = 0
                });
            }, ColumnOptions()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.Contains("Marker"));
        Assert.InRange(X(pdf, 2, "Marker001"), 39.9, kind == "list" ? 80 : 40.1);
        AssertMarkers(pdf, 1);
    }

    [Theory]
    [InlineData("heading")]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("panel")]
    public void Columns_KeepWithNextChainMovesTogetherPastUnusedPartialColumns(string kind) {
        byte[] bytes = PdfDocument.Create(Options()).Paragraph(p => p.Text(Markers("Prelude", 12)), style: ParagraphStyle())
            .Columns(columns => {
                if (kind == "heading") columns.H1("Marker001", style: new PdfHeadingStyle {
                    FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0
                });
                else if (kind == "paragraph") {
                    var style = ParagraphStyle(); style.KeepWithNext = true;
                    columns.Paragraph(p => p.Text("Marker001"), style: style);
                } else if (kind == "list") columns.RichNumbered(new[] { new PdfListItem("Marker001") }, style: new PdfListStyle {
                    LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true
                });
                else columns.Panel(content => content.Paragraph(p => p.Text("Marker001"), style: ParagraphStyle()), new PdfPanelStyle {
                    PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true
                });
                var body = ParagraphStyle(); body.KeepTogether = true;
                columns.Paragraph(p => p.Text(Markers("Body", 4)), style: body);
            }, ColumnOptions()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.Contains("Marker") || word.Text.Contains("Body"));
        double x = X(pdf, 2, "Body001");
        Assert.InRange(x, 39.9, 40.1);
        Assert.InRange(X(pdf, 2, "Marker001"), 39.9, kind == "list" ? 80 : 40.1);
        Assert.InRange(X(pdf, 2, "Body004"), x - .1, x + .1);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("panel")]
    [InlineData("flow")]
    public void Columns_KeptContentCannotRetryBeyondItsPaddedPhysicalCapacity(string kind) {
        var document = PdfDocument.Create(Options()).Panel(outer => {
            outer.Paragraph(p => p.Text("Prelude"), style: ParagraphStyle());
            outer.Columns(columns => AddKeptContent(columns, kind, 16), ColumnOptions());
        }, new PdfPanelStyle { PaddingX = 0, PaddingY = 20, SpacingBefore = 0, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("panel")]
    [InlineData("flow")]
    public void Columns_KeptContentMovesFromPartialColumnsToAFullPhysicalPage(string kind) {
        byte[] bytes = PdfDocument.Create(Options())
            .Paragraph(p => p.Text(Markers("Prelude", 12)), style: ParagraphStyle())
            .Columns(columns => AddKeptContent(columns, kind, 13), ColumnOptions()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.Contains("Marker"));
        for (int index = 1; index <= 13; index++)
            Assert.InRange(X(pdf, 2, $"Marker{index:D3}"), 39.9, kind == "list" ? 80 : 40.1);
        AssertMarkers(pdf, 13);
    }

    [Theory]
    [InlineData("heading")]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("panel")]
    [InlineData("flow")]
    public void Columns_BlockStartingANewPhysicalPageRetainsItsSiblingsForBalancing(string kind) {
        byte[] bytes = PdfDocument.Create(Options()).Columns(columns => {
            columns.Paragraph(p => p.Text(Markers("Prelude", 32)), style: ParagraphStyle());
            if (kind == "heading") columns.H1("Marker001", style: new PdfHeadingStyle {
                FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0
            });
            else AddKeptContent(columns, kind, 4);
            int count = kind == "heading" ? 9 : 6;
            columns.Paragraph(p => p.Text(Markers("Body", count)), style: ParagraphStyle());
        }, ColumnOptions()).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        int leftLast = kind == "heading" ? 4 : 1;
        Assert.InRange(X(pdf, 2, $"Body{leftLast:D3}"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, $"Body{leftLast + 1:D3}"), 259.9, 260.1);
        AssertMarkers(pdf, kind == "heading" ? 1 : 4);
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        Assert.Equal(Enumerable.Range(1, kind == "heading" ? 9 : 6), Matches(text, "Body"));
    }

    private static void AddKeptContent(PdfContentBuilder content, string kind, int count) {
        if (kind == "paragraph") {
            var style = ParagraphStyle(); style.KeepTogether = true;
            content.Paragraph(p => p.Text(Markers("Marker", count)), style: style);
        } else if (kind == "list") content.RichNumbered(Enumerable.Range(1, count)
            .Select(index => new PdfListItem($"Marker{index:D3}")), style: new PdfListStyle {
                LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true
            });
        else if (kind == "panel") content.Panel(panel => panel.Paragraph(p => p.Text(Markers("Marker", count)), style: ParagraphStyle()),
            new PdfPanelStyle { PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true });
        else content.Flow(flow => flow.Paragraph(p => p.Text(Markers("Marker", count)), style: ParagraphStyle()),
            new PdfFlowOptions { KeepTogether = true });
    }

    private static string Markers(string prefix, int count) => string.Join("\n", Enumerable.Range(1, count).Select(index => $"{prefix}{index:D3}"));
    private static IEnumerable<int> Matches(string text, string prefix) => System.Text.RegularExpressions.Regex.Matches(text, prefix + @"(\d{3})")
        .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value));
    private static void AssertMarkers(PdfPigDocument pdf, int count) => Assert.Equal(Enumerable.Range(1, count),
        Matches(string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text)), "Marker"));
    private static double X(PdfPigDocument pdf, int page, string marker) => pdf.GetPage(page).GetWords()
        .Single(word => word.Text == marker || word.Text.EndsWith("." + marker, StringComparison.Ordinal)).BoundingBox.Left;
    private static PdfOptions Options() => new() { PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
        MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, CompressContentStreams = false, MaxGeneratedPages = 4 };
    private static PdfParagraphStyle ParagraphStyle() => new() { LineSpacing = PdfLineSpacing.Exactly(20),
        SpacingBefore = 0, SpacingAfter = 0, WidowControl = false };
    private static PdfMultiColumnOptions ColumnOptions() => new() { Gap = 20, BalanceLastPage = true };
}
