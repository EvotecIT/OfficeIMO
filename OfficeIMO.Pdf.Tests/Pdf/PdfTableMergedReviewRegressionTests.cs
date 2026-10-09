using OfficeIMO.Pdf;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void TableMovedBeforeItsFirstRowDoesNotLeaveAnEmptyTaggedTable() {
        PdfTableStyle style = Style();
        style.MinimumBodyRowsOnFirstPage = 0;
        style.RowMinHeights = new() { 60, null };
        PdfDocument document = PdfDocument.Create(Options(140)).TaggedPdfCatalogMarkers();
        document.Compose(builder => builder.Page(page => page.Content(content => {
            content.Paragraph(p => p.Text("Intro1\nIntro2\nIntro3"), style: new PdfParagraphStyle {
                FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(14), SpacingBefore = 0,
                SpacingAfter = 0, WidowControl = false
            });
            content.Table(new[] {
                new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), PdfTableCell.TextCell("Ready") },
                new[] { PdfTableCell.TextCell("Done") }
            }, style: style);
        })));
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(1, Regex.Matches(Encoding.ASCII.GetString(bytes), @"/S /Table\b").Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergedCellFillDoesNotCoverTheDefaultGrid(bool inRow) {
        PdfTableStyle style = Style();
        style.MinRowHeight = 30;
        style.BorderWidth = 2;
        style.BorderColor = new PdfColor(.61, .21, .11);
        style.CellFills = new() { [(0, 0)] = new PdfColor(.31, .41, .51) };
        PdfDocument document = PdfDocument.Create(Options(200));
        AddTable(document, new[] {
            new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow);
        string raw = Encoding.ASCII.GetString(document.ToBytes());
        int fill = raw.LastIndexOf("0.31 0.41 0.51 rg", StringComparison.Ordinal);
        int border = raw.LastIndexOf("0.61 0.21 0.11 RG", StringComparison.Ordinal);
        Assert.True(fill >= 0 && border > fill, "The visible default grid must be painted after the opaque merged fill.");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MergedTailRespectsDisabledRowBreaks(bool inRow, bool perRow) {
        PdfTableStyle style = Style();
        style.AllowRowBreakAcrossPages = perRow;
        if (perRow) style.RowAllowBreakAcrossPages = new() { false, false, false };
        PdfDocument document = PdfDocument.Create(Options(140));
        AddTable(document, new[] {
            new[] { PdfTableCell.Merge(string.Join("\n", Enumerable.Range(1, 30).Select(n => $"L{n}")), rowSpan: 3), PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, inRow);
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void MergedTailContinuationReservesItsFirstTextLineBeforeOptionalSpacing() {
        string[] tokens = Enumerable.Range(1, 8).Select(n => $"Tall{n}").ToArray();
        PdfTextRun[] runs = { PdfTextRun.Normal(string.Join("\n", tokens)) };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(runs, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(60), widowControl: false)
        }, rowSpan: 3);
        PdfTableStyle style = Style();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.RowMinHeights = new() { 20, null, null, null };
        style.PageContinuationSpacingBefore = 60;
        PdfDocument document = PdfDocument.Create(Options(140));
        AddTable(document, new[] {
            new[] { PdfTableCell.TextCell("Header"), PdfTableCell.TextCell("Value") },
            new[] { anchor, PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        }, style, false);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        string[] words = pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text).ToArray();
        foreach (string token in tokens) Assert.Single(words, word => word == token);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergedTailReservesContainerClosingPaddingOnlyOnItsLastFragment(bool inRow) {
        string[] tokens = Enumerable.Range(1, 8).Select(n => $"Tall{n}").ToArray();
        PdfTextRun[] runs = { PdfTextRun.Normal(string.Join("\n", tokens) + "\nEnd") };
        PdfTextRun[] ending = { PdfTextRun.Normal("End") };
        PdfTextRun[] tall = { PdfTextRun.Normal(string.Join("\n", tokens)) };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(tall, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(90), widowControl: false),
            new PdfTableCellParagraph(ending, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(15), widowControl: false)
        }, rowSpan: 3);
        PdfDocument document = PdfDocument.Create(Options(160));
        document.Compose(builder => builder.Page(page => page.Content(content => content
            .Element(element => element.Style(new PdfPanelStyle {
                PaddingX = 0, PaddingY = 20, SpacingBefore = 0, SpacingAfter = 0
            }).Content(inner => {
                inner.Paragraph(p => p.Text("Prelude"), style: new PdfParagraphStyle {
                    FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(15),
                    SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
                });
                PdfTableCell[][] rows = {
                    new[] { anchor, PdfTableCell.TextCell("Ready") },
                    new[] { PdfTableCell.TextCell("Next") },
                    new[] { PdfTableCell.TextCell("Done") }
                };
                if (inRow) inner.Row(row => row.PercentColumn(100, column => column.Table(rows, style: Style())));
                else inner.Table(rows, style: Style());
            })) )));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        string[] words = pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text).ToArray();
        foreach (string token in tokens.Concat(new[] { "End" })) Assert.Single(words, word => word == token);
    }

    [Fact]
    public void DeferredMergedParagraphRetainsEveryLineWhenItsNeighborContinuesIntoANarrowerColumn() {
        PdfTextRun[] runs = { PdfTextRun.Normal("Kept1\nKept2\nKept3\nKept4") };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(runs, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(18),
                keepTogether: true, widowControl: false)
        }, rowSpan: 3);
        string[] neighbors = Enumerable.Range(1, 8).Select(n => $"N{n}").ToArray();
        PdfTableStyle style = Style();
        style.MinimumBodyRowsOnFirstPage = 0;
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 460, PageHeight = 140, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12
        }).Columns(content => {
            content.Paragraph(p => p.Text("Intro1\nIntro2\nIntro3"), style: new PdfParagraphStyle {
                FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(14), SpacingBefore = 0,
                SpacingAfter = 0, WidowControl = false
            });
            content.Table(new[] {
                new[] { anchor, PdfTableCell.TextCell(string.Join("\n", neighbors)) },
                new[] { PdfTableCell.TextCell("Next") },
                new[] { PdfTableCell.TextCell("Done") }
            }, style: style);
        }, new PdfMultiColumnOptions {
            ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20),
                new PdfFlowColumn(PdfColumnWidth.Fixed(100))
            }, BalanceLastPage = false
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        string[] words = pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text).ToArray();
        foreach (string token in Enumerable.Range(1, 4).Select(n => $"Kept{n}").Concat(neighbors))
            Assert.True(words.Count(word => word == token) == 1,
                $"Expected one {token}; extracted: {string.Join(", ", words)}");
    }
}
