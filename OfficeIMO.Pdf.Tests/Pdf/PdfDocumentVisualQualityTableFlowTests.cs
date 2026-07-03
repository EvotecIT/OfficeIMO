using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Table_WrapsLongCellTextInsideContentArea() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Area", "Status" },
                new[] {
                    "Generation",
                    "This is a long table cell value that should wrap instead of drawing across the next column or past the page margin."
                }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(rightMost, double.NegativeInfinity, contentRight + 1);

        int statusLineCount = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value) && letter.StartBaseLine.X > options.MarginLeft + 250)
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.True(statusLineCount > 1, "Expected the long table cell to wrap to multiple visual lines.");
    }

    [Fact]
    public void Table_UsesProportionalGlyphWidthsForWideCellWrapping() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "WWWWWWWW" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double cellRight = options.MarginLeft + 80;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, cellRight + 1);
    }

    [Fact]
    public void Table_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowCells() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { new string('i', 20) }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnTable_UsesProportionalGlyphWidthsForWideCellWrapping() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "WWWWWWWW" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double cellRight = options.MarginLeft + 80;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column table wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, cellRight + 1);
    }

    [Fact]
    public void RowColumnTable_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowCells() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { new string('i', 20) }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void List_UsesProportionalGlyphWidthsForWideBulletWrapping() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Bullets(new[] { "WWWWWWWW" })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected bullet-list wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void List_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowBullets() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Bullets(new[] { new string('i', 20) })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnList_UsesProportionalGlyphWidthsForWideNumberedWrapping() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Numbered(new[] { "WWWWWWWW" }))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column numbered-list wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void RowColumnList_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowNumberedItems() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Numbered(new[] { new string('i', 20) }))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void Table_BreaksLongUnspacedTokensAfterShortPrefixInsideContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Field", "Value" },
                new[] {
                    "Token",
                    "id " + new string('X', 72)
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(rightMost, double.NegativeInfinity, contentRight + 1);

        int tokenLineCount = page.Letters
            .Where(letter => letter.Value == "X")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.True(tokenLineCount > 2, "Expected the long unspaced token to split across multiple table cell lines.");
    }

    [Fact]
    public void Table_CellTextThatEscapesCellRectanglesIsClipped() {
        var style = TableStyles.Minimal();
        style.CellPaddingX = 8;
        style.CellPaddingY = 5;
        style.RowBaselineOffset = 40;

        byte[] topLevelBytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Long", "This table cell deliberately wraps so the writer has to emit more than one clipped text line." }
            }, style: style)
            .ToBytes();

        string topLevelContent = string.Join("\n", GetPageContentStreams(topLevelBytes, 1));
        int topLevelClipCount = Regex.Matches(topLevelContent, " re W n\\nBT\\n[\\s\\S]{0,160}?/F").Count;
        Assert.True(topLevelClipCount >= 4, "Expected top-level table cell text to be clipped by PDF cell rectangles.");

        byte[] rowColumnBytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Value" },
                                    new[] { "Long", "Column-local table cells also get clipped to the cell content rectangle." }
                                }, style: style))))))
            .ToBytes();

        string rowColumnContent = string.Join("\n", GetPageContentStreams(rowColumnBytes, 1));
        int rowColumnClipCount = Regex.Matches(rowColumnContent, " re W n\\nBT\\n[\\s\\S]{0,160}?/F").Count;
        Assert.True(rowColumnClipCount >= 4, "Expected row-column table cell text to be clipped by PDF cell rectangles.");
    }

    [Fact]
    public void Table_PaginatesLongTablesAndRepeatsHeaderRows() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        var rows = new List<string[]> {
            new[] { "Metric", "Status" }
        };
        for (int i = 1; i <= 28; i++) {
            rows.Add(new[] { "Item " + i.ToString(), "Completed without clipping" });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Metric", page.Text);
            Assert.Contains("Status", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected table text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("Item 1", pdf.GetPage(1).Text);
        Assert.Contains("Item 28", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_RepeatsConfiguredHeaderRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.ColumnWidthWeights = new List<double> { 1, 1 };

        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int i = 1; i <= 30; i++) {
            rows.Add(new[] { "Check " + i.ToString(), "Team " + i.ToString() });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the two-row header table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var pageText = pdf.GetPage(pageNumber).Text;
            Assert.Contains("Group", pageText);
            Assert.Contains("State", pageText);
            Assert.Contains("Metric", pageText);
            Assert.Contains("Owner", pageText);
        }

        Assert.Contains("Check 1", pdf.GetPage(1).Text);
        Assert.Contains("Check 30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_CanStyleHeaderRowsWithoutRepeatingThemAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 0;
        style.ColumnWidthWeights = new List<double> { 1, 1 };

        var rows = new List<string[]> {
            new[] { "VisualHdr", "State" }
        };
        for (int i = 1; i <= 30; i++) {
            rows.Add(new[] { "StyledOnly " + i.ToString(CultureInfo.InvariantCulture), "Ready" });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the visually styled header table to continue onto another page.");

        int headerOccurrences = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "VisualHdr");
        Assert.Equal(1, headerOccurrences);
        Assert.Contains("StyledOnly 30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_RepeatsHeaderRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 210,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.ColumnWidthWeights = new List<double> { 1, 1 };

        var rows = new List<string[]> {
            new[] { "ColMetric", "ColValue" }
        };
        for (int i = 1; i <= 28; i++) {
            rows.Add(new[] { "ColumnCheck " + i.ToString(CultureInfo.InvariantCulture), "Ready" });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the column-local table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var pageText = pdf.GetPage(pageNumber).Text;
            Assert.Contains("ColMetric", pageText);
            Assert.Contains("ColValue", pageText);
        }

        Assert.Contains("ColumnCheck 1", pdf.GetPage(1).Text);
        Assert.Contains("ColumnCheck 28", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_RendersConfiguredFooterRowsAtEndOfLongTables() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.FooterRowCount = 1;
        style.FooterFill = PdfColor.FromRgb(230, 230, 230);
        style.FooterTextColor = PdfColor.FromRgb(20, 20, 20);

        var rows = new List<string[]> {
            new[] { "Metric", "Value" }
        };
        for (int i = 1; i <= 30; i++) {
            rows.Add(new[] { "Item " + i.ToString(), i.ToString() });
        }
        rows.Add(new[] { "Total", "30" });

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long table with footer rows to continue onto another page.");
        Assert.DoesNotContain("Total", pdf.GetPage(1).Text);
        Assert.Contains("Total", pdf.GetPage(pdf.NumberOfPages).Text);
        Assert.Contains("30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_KeepTogetherMovesWholeTableToNextPage() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepTogether = true;

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 52
            })
            .Table(new[] {
                new[] { "KeepA", "Ready" },
                new[] { "KeepB", "Ready" },
                new[] { "KeepC", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepA", pdf.GetPage(1).Text);
        Assert.Contains("KeepA", pdf.GetPage(2).Text);
        Assert.Contains("KeepC", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnTable_KeepTogetherMovesWholeTableToNextPage() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepTogether = true;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => {
                                column.Paragraph(p => p.Text("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = 52
                                });
                                column.Table(new[] {
                                    new[] { "ColumnKeepA", "Ready" },
                                    new[] { "ColumnKeepB", "Ready" },
                                    new[] { "ColumnKeepC", "Ready" }
                                }, style: style);
                            })))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ColumnIntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepA", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepA", pdf.GetPage(2).Text);
        Assert.Contains("ColumnKeepC", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Table_KeepTogetherRejectsTableTallerThanContentArea() {
        var style = TableStyles.Minimal();
        style.KeepTogether = true;
        style.HeaderRowCount = 0;
        style.LineHeight = 2.0;

        var rows = Enumerable.Range(1, 10)
            .Select(i => new[] { "KeepTooTall" + i.ToString(CultureInfo.InvariantCulture) })
            .ToArray();

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 260,
                    PageHeight = 160,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .Table(rows, style: style)
                .ToBytes());

        Assert.Contains("Table height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }


}
