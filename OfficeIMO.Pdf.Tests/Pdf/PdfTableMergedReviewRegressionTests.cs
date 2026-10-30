using OfficeIMO.Pdf;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void TableMovedBeforeItsFirstRowDoesNotLeaveAnEmptyTaggedTable() {
        PdfTableStyle style = Style();
        style.MinimumBodyRowsOnFirstPage = 0;
        style.RowMinHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(60), null };
        PdfDocument document = PdfDocument.Create(Options(140)).TaggedPdfCatalogMarkers();
        document.Compose(builder => builder.Page(page => page.Content(content => {
            content.Paragraph(p => p.Text("Intro1\nIntro2\nIntro3"), style: new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(14)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false
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
        style.MinRowHeight = global::OfficeIMO.OfficeLength.Points(30);
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(2);
        style.BorderColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.61, .21, .11);
        style.CellFills = new() { [(0, 0)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.31, .41, .51) };
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
    public void LaterMergedFillPreservesEarlierExplicitNeighborBorder(bool inRow, bool mergedNeighbor) {
        PdfTableStyle style = Style();
        style.MinRowHeight = global::OfficeIMO.OfficeLength.Points(30);
        style.BorderColor = null;
        style.CellBorders = new() { [(0, 0)] = new PdfCellBorder {
            Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.61, .21, .11), Width = global::OfficeIMO.OfficeLength.Points(4),
            Top = false, Bottom = false, Left = false, Right = true
        } };
        style.CellFills = new() { [(0, 1)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.71, .81, .61) };
        PdfDocument document = PdfDocument.Create(Options(200));
        AddTable(document, new[] {
            new[] { mergedNeighbor ? PdfTableCell.Merge("Alpha", rowSpan: 2) : PdfTableCell.TextCell("Alpha"),
                PdfTableCell.Merge("Beta", rowSpan: 3), PdfTableCell.TextCell("Ready") },
            mergedNeighbor ? new[] { PdfTableCell.TextCell("Next") } :
                new[] { PdfTableCell.TextCell("Middle"), PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Tail"), PdfTableCell.TextCell("Done") }
        }, style, inRow);
        string raw = Encoding.ASCII.GetString(document.ToBytes());
        int fill = raw.LastIndexOf("0.71 0.81 0.61 rg", StringComparison.Ordinal);
        int border = raw.LastIndexOf("0.61 0.21 0.11 RG", StringComparison.Ordinal);
        Assert.True(fill >= 0 && border > fill,
            "A later opaque merged fill must preserve the neighboring explicit border's full width and color.");
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
            new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(60)), widowControl: false)
        }, rowSpan: 3);
        PdfTableStyle style = Style();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.RowMinHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(20), null, null, null };
        style.PageContinuationSpacingBefore = global::OfficeIMO.OfficeLength.Points(60);
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
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MergedTailContinuationReservesItsParagraphBoundaryBeforeOptionalSpacing(bool oversizedWidowParagraph, bool inClosingPanelRow) {
        string[] tallTokens = Enumerable.Range(1, 8).Select(n => $"Tall{n}").ToArray();
        string[] keptTokens = Enumerable.Range(1, oversizedWidowParagraph ? 3 : 4).Select(n => $"Kept{n}").ToArray();
        PdfTextRun[] tall = { PdfTextRun.Normal(string.Join("\n", tallTokens)) };
        PdfTextRun[] kept = { PdfTextRun.Normal(string.Join("\n", keptTokens)) };
        PdfTextRun[] runs = { PdfTextRun.Normal(string.Join("\n", tallTokens.Concat(keptTokens))) };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(tall, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(inClosingPanelRow ? 80 : 60)), widowControl: false),
            new PdfTableCellParagraph(kept, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(oversizedWidowParagraph ? 30 : 18)),
                keepTogether: !oversizedWidowParagraph, widowControl: oversizedWidowParagraph)
        }, rowSpan: 3);
        PdfTableStyle style = Style();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.RowMinHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(20), null, null, null };
        style.PageContinuationSpacingBefore = global::OfficeIMO.OfficeLength.Points(60);
        PdfDocument document = PdfDocument.Create(Options(140));
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.TextCell("Header"), PdfTableCell.TextCell("Value") },
            new[] { anchor, PdfTableCell.TextCell("Ready") },
            new[] { PdfTableCell.TextCell("Next") },
            new[] { PdfTableCell.TextCell("Done") }
        };
        if (inClosingPanelRow) {
            document.Compose(builder => builder.Page(page => page.Content(content => content
                .Element(element => element.Style(new PdfPanelStyle {
                    PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(20), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
                    RepeatFragmentDecoration = false
                }).Content(inner => inner.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))))));
        } else AddTable(document, rows, style, false);
        using CancellationTokenSource cancellation = new(TimeSpan.FromSeconds(5));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes(cancellationToken: cancellation.Token));
        foreach (string token in tallTokens.Concat(keptTokens))
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == token);
        // Closing padding makes this paragraph taller than the final usable frame,
        // so its keep rule may relax there while every line must still be preserved.
        if (!oversizedWidowParagraph && !inClosingPanelRow)
            Assert.Single(pdf.GetPages(), page => keptTokens.All(token => page.GetWords().Any(word => word.Text == token)));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void MergedTailReservesContainerClosingPaddingOnlyOnItsLastFragment(bool inRow, bool repeatDecoration) {
        string[] tokens = Enumerable.Range(1, 8).Select(n => $"Tall{n}").ToArray();
        PdfTextRun[] runs = { PdfTextRun.Normal(string.Join("\n", tokens) + "\nEnd") };
        PdfTextRun[] ending = { PdfTextRun.Normal("End") };
        PdfTextRun[] tall = { PdfTextRun.Normal(string.Join("\n", tokens)) };
        PdfTableCell anchor = new(runs, new[] {
            new PdfTableCellParagraph(tall, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(90)), widowControl: false),
            new PdfTableCellParagraph(ending, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(15)), widowControl: false)
        }, rowSpan: 3);
        PdfDocument document = PdfDocument.Create(Options(160));
        document.Compose(builder => builder.Page(page => page.Content(content => content
            .Element(element => element.Style(new PdfPanelStyle {
                PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(20), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
                RepeatFragmentDecoration = repeatDecoration
            }).Content(inner => {
                inner.Paragraph(p => p.Text("Prelude"), style: new PdfParagraphStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(12), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(15)),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false
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
            new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)),
                keepTogether: true, widowControl: false)
        }, rowSpan: 3);
        string[] neighbors = Enumerable.Range(1, 8).Select(n => $"N{n}").ToArray();
        PdfTableStyle style = Style();
        style.MinimumBodyRowsOnFirstPage = 0;
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(460), PageHeight = global::OfficeIMO.OfficeLength.Points(140), MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12)
        }).Columns(content => {
            content.Paragraph(p => p.Text("Intro1\nIntro2\nIntro3"), style: new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(14)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false
            });
            content.Table(new[] {
                new[] { anchor, PdfTableCell.TextCell(string.Join("\n", neighbors)) },
                new[] { PdfTableCell.TextCell("Next") },
                new[] { PdfTableCell.TextCell("Done") }
            }, style: style);
        }, new PdfMultiColumnOptions {
            ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(300)), global::OfficeIMO.OfficeLength.Points(20)),
                new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(100)))
            }, BalanceLastPage = false
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        string[] words = pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text).ToArray();
        foreach (string token in Enumerable.Range(1, 4).Select(n => $"Kept{n}").Concat(neighbors))
            Assert.True(words.Count(word => word == token) == 1,
                $"Expected one {token}; extracted: {string.Join(", ", words)}");
    }
}
