using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MergedAnchorRetainsItsFirstTaggedColumn(bool inRow, bool linked) {
        PdfTableCell[][] rows = {
            new[] { new PdfTableCell("Alpha", rowSpan: 2, linkUri: linked ? "https://example.com" : null), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Done") }
        };
        PdfDocument document = PdfDocument.Create(Options(200)).TaggedPdfCatalogMarkers();
        AddTable(document, rows, Style(), inRow);
        string raw = Encoding.ASCII.GetString(document.ToBytes());
        Dictionary<string, string> objects = Regex.Matches(raw, @"(\d+) 0 obj\s*(.*?)\s*endobj", RegexOptions.Singleline)
            .Cast<Match>()
            .ToDictionary(match => match.Groups[1].Value, match => match.Groups[2].Value);
        string firstRow = objects.Values.First(value => value.Contains("/S /TR", StringComparison.Ordinal));
        string firstCell = Regex.Match(firstRow, @"/K\s*\[\s*(\d+) 0 R").Groups[1].Value;
        Assert.Contains("/RowSpan 2", objects[firstCell], StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlexibleMergedFillUsesTheTableRoundedPerimeter(bool inRow) {
        PdfTableStyle style = Style();
        style.MinRowHeight = 30;
        style.CornerRadius = 10;
        style.CellFills = new() { [(0, 0)] = new PdfColor(.31, .41, .51) };
        style.CellBorders = new() { [(0, 0)] = new PdfCellBorder { Color = new PdfColor(.61, .21, .11), Width = 2 } };
        PdfDocument document = PdfDocument.Create(Options(200));
        AddTable(document, new[] {
            new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow);
        string raw = Encoding.ASCII.GetString(document.ToBytes());
        int fill = raw.IndexOf("0.31 0.41 0.51 rg", StringComparison.Ordinal);
        Assert.True(fill >= 0);
        int clip = raw.LastIndexOf("W n", fill, StringComparison.Ordinal);
        Assert.True(clip >= 0, "The merged fill needs a rounded clipping perimeter.");
        int save = raw.LastIndexOf("q\n", clip, StringComparison.Ordinal);
        Assert.Contains("W n", raw.Substring(save, fill - save), StringComparison.Ordinal);
        Assert.Contains(" c", raw.Substring(save, fill - save), StringComparison.Ordinal);
        Assert.DoesNotContain("Q\n", raw.Substring(clip, fill - clip), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    public void OneRowMergedFragmentsKeepHiddenVerticalSegments(bool inRow, int hiddenRow) {
        PdfTableStyle style = Style();
        style.MinRowHeight = 60;
        style.CellBorders = new() { [(0, 0)] = new PdfCellBorder {
            Color = new PdfColor(.61, .21, .11), Width = 2,
            Top = false, Bottom = false, Left = false, Right = true,
            HiddenRightRowSegments = new() { hiddenRow }
        } };
        PdfDocument document = PdfDocument.Create(Options(120));
        AddTable(document, new[] {
            new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow);
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        string content = string.Join("\n", GetPageContentStreams(bytes, hiddenRow + 1));
        Assert.DoesNotContain("0.61 0.21 0.11 RG", content, StringComparison.Ordinal);
        string visible = string.Join("\n", GetPageContentStreams(bytes, 2 - hiddenRow));
        Assert.Contains("0.61 0.21 0.11 RG", visible, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, "together")]
    [InlineData(true, "together")]
    [InlineData(false, "next")]
    [InlineData(true, "next")]
    [InlineData(false, "widow")]
    [InlineData(true, "widow")]
    public void MergedFragmentsRespectParagraphBoundaries(bool inRow, string rule) {
        string firstText = rule == "next" ? "First1\nFirst2" : rule == "widow" ? "First1\nFirst2\nFirst3" : "First1\nFirst2\nFirst3\nFirst4";
        string secondText = string.Join("\n", Enumerable.Range(1, 8 - firstText.Split('\n').Length).Select(n => $"Second{n}"));
        PdfTextRun[] first = { PdfTextRun.Normal(firstText) };
        PdfTextRun[] second = { PdfTextRun.Normal(secondText) };
        PdfTableCell cell = new(first.Concat(second), new[] {
            new PdfTableCellParagraph(first, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20),
                keepTogether: rule == "together", keepWithNext: rule == "next", widowControl: rule == "widow"),
            new PdfTableCellParagraph(second, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false)
        }, rowSpan: 3);
        PdfDocument document = PdfDocument.Create(Options(140));
        AddTable(document, new[] {
            new[] { cell, PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        }, Style(), inRow);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, word.Text))).ToArray();
        int PageOf(string marker) => Assert.Single(words, word => word.Text == marker).Number;
        if (rule == "next") Assert.Equal(PageOf("First2"), PageOf("Second1"));
        else Assert.Equal(PageOf("First1"), PageOf(rule == "widow" ? "First3" : "First4"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergedBorderMasksUseTheAdmittedPartialRowHeight(bool inRow) {
        PdfTableStyle style = Style();
        style.LineHeight = 1;
        style.RowMinHeights = new() { null, 30 };
        style.CellBorders = new() { [(0, 0)] = new PdfCellBorder {
            Color = new PdfColor(.61, .21, .11), Width = 2,
            Top = false, Bottom = false, Left = false, Right = true,
            HiddenRightRowSegments = new() { 0 }
        } };
        PdfDocument document = PdfDocument.Create(Options(140));
        AddTable(document, new[] {
            new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), PdfTableCell.TextCell(string.Join("\n", Enumerable.Range(1, 10).Select(n => $"Long{n}"))) },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow);
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        string content = string.Join("\n", GetPageContentStreams(bytes, 2));
        Match stroke = Regex.Match(content, @"0\.61 0\.21 0\.11 RG\s+.*?([\d.]+) ([\d.]+) m\s+([\d.]+) ([\d.]+) l", RegexOptions.Singleline);
        Assert.True(stroke.Success, "The next physical row owns a visible border after the hidden partial row.");
        double top = double.Parse(stroke.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
        double bottom = double.Parse(stroke.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
        Assert.InRange(top - bottom, 29.9D, 30.1D);
    }

    [Fact]
    public void MergedTailReflowsAcrossUnequalSequentialColumns() {
        string[] tokens = Enumerable.Range(1, 200).Select(n => $"Word{n:D3}").ToArray();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 460, PageHeight = 140, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12
        }).Columns(content => content.Table(new[] {
            new[] { PdfTableCell.Merge(string.Join(" ", tokens), rowSpan: 3), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style: Style()), new PdfMultiColumnOptions {
            ColumnDefinitions = new[] { new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(100)) },
            BalanceLastPage = false
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text).ToArray();
        foreach (string token in tokens) Assert.Single(words, word => word == token);
    }

    [Theory]
    [InlineData(false, 4, 0D)]
    [InlineData(true, 4, 0D)]
    [InlineData(false, 5, 0D)]
    [InlineData(true, 5, 0D)]
    [InlineData(false, 4, 20D)]
    public void MergedParagraphKeepUsesCapacityAfterTheRepeatedHeader(bool inRow, int lineCount, double spacingAfter) {
        PdfTextRun[] runs = { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, lineCount).Select(n => $"Kept{n}"))) };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(runs, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(18), keepTogether: true, widowControl: false)
        }, rowSpan: 3);
        PdfTableStyle style = Style();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.SpacingAfter = spacingAfter;
        style.RowMinHeights = new() { 20, null, null, null };
        PdfDocument document = PdfDocument.Create(Options(140));
        AddTable(document, new[] {
            new[] { PdfTableCell.TextCell("Header"), PdfTableCell.TextCell("Value") },
            new[] { anchor, PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow, reserveClosing: spacingAfter > 0D);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, word.Text))).ToArray();
        int[] pages = Enumerable.Range(1, lineCount).Select(n => Assert.Single(words, word => word.Text == $"Kept{n}").Number).ToArray();
        if (lineCount * 18 <= 80 - spacingAfter) Assert.Single(pages.Distinct());
        else Assert.True(pages.Distinct().Count() > 1, "An oversized kept paragraph must relax within the usable body capacity.");
    }

    private static PdfOptions Options(double height) => new() {
        PageWidth = 300, PageHeight = height, MarginLeft = 20, MarginRight = 20,
        MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, CompressContentStreams = false
    };

    private static PdfTableStyle Style() => new() {
        HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12,
        MinRowHeight = 0, SpacingBefore = 0, SpacingAfter = 0, BorderWidth = 0, RowSeparatorWidth = 0
    };

    private static void AddTable(PdfDocument document, PdfTableCell[][] rows, PdfTableStyle style, bool inRow, bool reserveClosing = false) {
        document.Compose(builder => builder.Page(page => page.Content(content => {
            if (reserveClosing) content.Element(element => element
                .Style(new PdfPanelStyle { PaddingX = 0, PaddingY = 10, SpacingBefore = 0, SpacingAfter = 0 })
                .Content(inner => inner.Table(rows, style: style)));
            else if (inRow) content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
            else content.Table(rows, style: style);
        })));
    }
}
