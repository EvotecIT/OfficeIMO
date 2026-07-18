using System.Globalization;
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

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Open(bytes).Read.TextBlocks();
        Assert.Single(blocks, block => block.Text.Contains("Left one", StringComparison.Ordinal));
        Assert.Single(blocks, block => block.Text.Contains("Right one", StringComparison.Ordinal));
        string raw = PdfEncoding.Latin1GetString(bytes);
        Assert.Contains("1 0 0 1 72 ", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 318 ", raw, StringComparison.Ordinal);
        Assert.Single(PdfReadDocument.Open(bytes).Pages);
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
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Container title", text, StringComparison.Ordinal);
        Assert.Contains("Container body", text, StringComparison.Ordinal);
        Assert.Contains("0.9 0.95 1 rg", raw, StringComparison.Ordinal);
        Assert.Contains("1.5 w", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Container_PaginatesAndDecoratesEveryPageFragment() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20,
            DefaultFontSize = 10,
            CompressContentStreams = false
        });

        document.Container(content => {
            for (int index = 1; index <= 24; index++) {
                int marker = index;
                content.Paragraph(paragraph => paragraph.Text("Decorated fragment line " + marker));
            }
        }, new PanelStyle {
            Background = new PdfColor(0.9D, 0.95D, 1D),
            BorderColor = new PdfColor(0.1D, 0.2D, 0.4D),
            BorderWidth = 1.5D,
            PaddingX = 8D,
            PaddingY = 6D,
            KeepTogether = false
        });

        byte[] bytes = document.ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        string raw = PdfEncoding.Latin1GetString(bytes);

        Assert.True(read.Pages.Count >= 2);
        Assert.Contains("Decorated fragment line 1", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Decorated fragment line 24", read.ExtractText(), StringComparison.Ordinal);
        Assert.True(CountOccurrences(raw, "0.9 0.95 1 rg") >= read.Pages.Count);
        Assert.True(CountOccurrences(raw, "1.5 w") >= read.Pages.Count);
    }

    [Fact]
    public void Container_RejectsUnsupportedNestedPageBreak() {
        PdfDocument document = PdfDocument.Create().Container(content => content.PageBreak());
        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    [Fact]
    public void Columns_BalancesOneLongParagraphAtWrappedLineBoundaries() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 40,
                DefaultFontSize = 10,
                CompressContentStreams = false
            })
            .Columns(columns => columns.Paragraph(paragraph => paragraph
                .Text("ColumnLine01").LineBreak()
                .Text("ColumnLine02").LineBreak()
                .Text("ColumnLine03").LineBreak()
                .Text("ColumnLine04").LineBreak()
                .Text("ColumnLine05").LineBreak()
                .Text("ColumnLine06").LineBreak()
                .Text("ColumnLine07").LineBreak()
                .Text("ColumnLine08")), new PdfMultiColumnOptions {
                    ColumnCount = 2,
                    Gap = 20,
                    BalanceLastPage = true,
                    BalanceParagraphLines = true
                })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        Assert.Single(read.Pages);
        Assert.Contains("ColumnLine01", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("ColumnLine08", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 40 ", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 220 ", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Columns_DefaultKeepTogetherParagraphIsNotSplitByLineBalancing() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 40,
                DefaultFontSize = 10,
                DefaultParagraphStyle = new PdfParagraphStyle {
                    KeepTogether = true
                },
                CompressContentStreams = false
            })
            .Columns(columns => columns.Paragraph(paragraph => paragraph
                .Text("KeepLine01").LineBreak()
                .Text("KeepLine02").LineBreak()
                .Text("KeepLine03").LineBreak()
                .Text("KeepLine04").LineBreak()
                .Text("KeepLine05").LineBreak()
                .Text("KeepLine06").LineBreak()
                .Text("KeepLine07").LineBreak()
                .Text("KeepLine08")), new PdfMultiColumnOptions {
                    ColumnCount = 2,
                    Gap = 20,
                    BalanceLastPage = true,
                    BalanceParagraphLines = true
                })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        PdfReadDocument read = PdfReadDocument.Open(bytes);

        Assert.Single(read.Pages);
        Assert.Contains("KeepLine01", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("KeepLine08", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 40 ", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 1 220 ", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Columns_BalancedRichLinesPreserveSpaceBeforeInlineElements() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 40,
                DefaultFontSize = 10,
                CompressContentStreams = false
            })
            .Columns(columns => columns.Paragraph(paragraph => paragraph
                .Text("ColumnLine01 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine02 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine03 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine04 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine05 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine06 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine07 ").InlineBox(8, 8).Text("after").LineBreak()
                .Text("ColumnLine08 ").InlineBox(8, 8).Text("after")), new PdfMultiColumnOptions {
                    ColumnCount = 2,
                    Gap = 20,
                    BalanceLastPage = true,
                    BalanceParagraphLines = true
                })
            .ToBytes();

        string text = PdfReadDocument.Open(bytes).ExtractText();

        for (int index = 1; index <= 8; index++) {
            string marker = "ColumnLine" + index.ToString("D2") + " after";
            Assert.Contains(marker, text, StringComparison.Ordinal);
            Assert.DoesNotContain(marker.Replace(" ", string.Empty), text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Columns_BalancedJustifiedParagraphPreservesSoftWrapJustification() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 40,
                DefaultFontSize = 10,
                CompressContentStreams = false
            })
            .Columns(columns => columns.Paragraph(paragraph => paragraph.Text(
                "Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee zulu " +
                "amber bronze copper denim emerald fuchsia garnet hazel ivory jade khaki lilac magenta navy ochre peach quartz ruby silver teal umber violet wheat yellow zinc"),
                align: PdfAlign.Justify), new PdfMultiColumnOptions {
                    ColumnCount = 2,
                    Gap = 20,
                    BalanceLastPage = true,
                    BalanceParagraphLines = true
                })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        double[] wordSpacings = raw
            .Split('\n')
            .Select(line => line.Trim())
            .Where(line => line.EndsWith(" Tw", StringComparison.Ordinal))
            .Select(line => double.TryParse(
                line.Substring(0, line.Length - 3),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double spacing) ? spacing : 0D)
            .ToArray();

        Assert.Contains(wordSpacings, spacing => spacing > 0.001D);
        Assert.Contains("1 0 0 1 40 ", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 220 ", raw, StringComparison.Ordinal);
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(token, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += token.Length;
        }

        return count;
    }
}
