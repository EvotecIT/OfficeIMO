using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfContainerAndColumnTests {
    [Fact]
    public void Columns_DistributesBlocksAcrossEqualFramesAndSupportsExplicitBreak() {
        byte[] bytes = PdfDocument.Create()
            .Columns(columns => {
                columns.Paragraph(paragraph => paragraph.Text("Left one"));
                columns.Paragraph(paragraph => paragraph.Text("Left two"));
                columns.ColumnBreak();
                columns.Paragraph(paragraph => paragraph.Text("Right one"));
            }, new PdfMultiColumnOptions { ColumnCount = 2, Gap = 24, BalanceLastPage = false })
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Load(bytes).Read.TextBlocks();
        Assert.Single(blocks, block => block.Text.Contains("Left one", StringComparison.Ordinal));
        Assert.Single(blocks, block => block.Text.Contains("Right one", StringComparison.Ordinal));
        string raw = PdfEncoding.Latin1GetString(bytes);
        Assert.Contains("1 0 0 1 72 ", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 318 ", raw, StringComparison.Ordinal);
        Assert.Single(PdfReadDocument.Load(bytes).Pages);
    }

    [Fact]
    public void Container_RendersNestedBlocksWithPaddingBackgroundAndBorder() {
        byte[] bytes = PdfDocument.Create()
            .Container(content => {
                content.H2("Container title");
                content.Paragraph(paragraph => paragraph.Text("Container body"));
            }, new PanelStyle {
                Background = new PdfColor(0.9D, 0.95D, 1D),
                BorderColor = new PdfColor(0.1D, 0.2D, 0.4D),
                BorderWidth = 1.5D,
                PaddingX = 12D,
                PaddingY = 10D,
                MaxWidth = 300D,
                Align = PdfAlign.Center
            })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();
        Assert.Contains("Container title", text, StringComparison.Ordinal);
        Assert.Contains("Container body", text, StringComparison.Ordinal);
        Assert.Contains("0.9 0.95 1 rg", raw, StringComparison.Ordinal);
        Assert.Contains("1.5 w", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Container_RejectsUnsupportedNestedPageBreak() {
        PdfDocument document = PdfDocument.Create().Container(content => content.PageBreak());
        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }
}
