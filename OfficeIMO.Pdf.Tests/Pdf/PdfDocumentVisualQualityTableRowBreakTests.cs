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
    public void Table_SplitsSingleTallRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Type", "Description" },
                new[] { "Finding", longValue }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very tall table row to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Type", page.Text);
            Assert.Contains("Description", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split row text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_SplitRowsUseMeasuredRichLineHeights() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        var richRuns = new[] {
            TextRun.Normal(string.Join(" ", Enumerable.Range(1, 30).Select(i => "large" + i.ToString("00", CultureInfo.InvariantCulture))), fontSize: 22)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { PdfTableCell.TextCell("Type"), PdfTableCell.TextCell("Description") },
                new[] { PdfTableCell.TextCell("Finding"), PdfTableCell.RichTextCell(richRuns) }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the rich table row to split across pages.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split rich table row text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("large01", pdf.GetPage(1).Text);
        Assert.Contains("large30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_SplitRowsUseMeasuredRichLineHeights() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        var richRuns = new[] {
            TextRun.Normal(string.Join(" ", Enumerable.Range(1, 30).Select(i => "large" + i.ToString("00", CultureInfo.InvariantCulture))), fontSize: 22)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.TextCell("Type"), PdfTableCell.TextCell("Description") },
                                    new[] { PdfTableCell.TextCell("Finding"), PdfTableCell.RichTextCell(richRuns) }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the rich row-column table row to split across pages.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split rich row-column table text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("large01", pdf.GetPage(1).Text);
        Assert.Contains("large30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_SplitsSingleTallRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Type", "Description" },
                                    new[] { "Finding", longValue }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very tall row-column table row to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Type", page.Text);
            Assert.Contains("Description", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split row-column row text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_SplitsAllowedMultiLineRowsIntoRemainingPageSpace() {
        var options = CreateRowSplitRemainderOptions();
        var style = CreateRowSplitRemainderStyle();
        var rows = CreateRowSplitRemainderRows();

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        AssertRowSplitUsesPageRemainder(bytes);
    }

    [Fact]
    public void Table_PageContinuationSpacingBefore_OffsetsContinuationPages() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        var rows = new List<string[]>();
        rows.Add(new[] { "HdrA", "HdrB" });
        for (int row = 1; row <= 28; row++) {
            rows.Add(new[] { "Row" + row.ToString("D2", CultureInfo.InvariantCulture), "Value" + row.ToString("D2", CultureInfo.InvariantCulture) });
        }

        var normalStyle = TableStyles.Minimal();
        normalStyle.ColumnWidthPoints = new List<double?> { 70, null };

        var spacedStyle = normalStyle.Clone();
        spacedStyle.PageContinuationSpacingBefore = 18D;

        byte[] normalBytes = PdfDocument.Create(options)
            .Table(rows, style: normalStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Table(rows, style: spacedStyle)
            .ToBytes();

        using var normalPdf = PdfPigDocument.Open(new MemoryStream(normalBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
        Assert.True(normalPdf.NumberOfPages > 1);
        Assert.True(spacedPdf.NumberOfPages > 1);

        double normalHeaderTop = GetWordTop(normalPdf, 2, "HdrA");
        double spacedHeaderTop = GetWordTop(spacedPdf, 2, "HdrA");
        Assert.True(spacedHeaderTop <= normalHeaderTop - 14D, $"Expected continuation spacing to lower page 2 header from {normalHeaderTop.ToString(CultureInfo.InvariantCulture)}; actual {spacedHeaderTop.ToString(CultureInfo.InvariantCulture)}.");
    }

    [Fact]
    public void Table_KeepsFinalTwoBodyRowsTogetherAtPageBreak() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.MinRowHeight = 18D;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        byte[] bytes = PdfDocument.Create(options)
            .Spacer(88)
            .Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Row01", "Alpha" },
                new[] { "Row02", "Beta" },
                new[] { "Row03", "Gamma" },
                new[] { "Row04", "Delta" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Row01", pdf.GetPage(1).Text);
        Assert.Contains("Row02", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Row03", pdf.GetPage(1).Text);
        Assert.Contains("Row03", pdf.GetPage(2).Text);
        Assert.Contains("Row04", pdf.GetPage(2).Text);
        Assert.Contains("Name", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnTable_KeepsFinalTwoBodyRowsTogetherAtPageBreak() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.MinRowHeight = 18D;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Spacer(94)
                                .Table(new[] {
                                    new[] { "Name", "Value" },
                                    new[] { "Row01", "Alpha" },
                                    new[] { "Row02", "Beta" },
                                    new[] { "Row03", "Gamma" },
                                    new[] { "Row04", "Delta" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Row01", pdf.GetPage(1).Text);
        Assert.Contains("Row02", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Row03", pdf.GetPage(1).Text);
        Assert.Contains("Row03", pdf.GetPage(2).Text);
        Assert.Contains("Row04", pdf.GetPage(2).Text);
        Assert.Contains("Name", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnTable_SplitsAllowedMultiLineRowsIntoRemainingPageSpace() {
        var options = CreateRowSplitRemainderOptions();
        var style = CreateRowSplitRemainderStyle();
        var rows = CreateRowSplitRemainderRows();

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(rows, style: style))))))
            .ToBytes();

        AssertRowSplitUsesPageRemainder(bytes);
    }

    [Fact]
    public void Table_DisallowRowBreakRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Table(new[] {
                    new[] { "Type", "Description" },
                    new[] { "Finding", longValue }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RowBreakPolicyAllowsSingleTallRows() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.RowAllowBreakAcrossPages = new List<bool?> { null, true };
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Type", "Description" },
                new[] { "Finding", longValue }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the per-row break policy to allow the tall row to split.");
        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_RowBreakPolicyRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = true;
        style.RowAllowBreakAcrossPages = new List<bool?> { null, false };
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Table(new[] {
                    new[] { "Type", "Description" },
                    new[] { "Finding", longValue }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RowBreakPolicyAllowsSingleTallRows() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.RowAllowBreakAcrossPages = new List<bool?> { null, true };
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Type", "Description" },
                                    new[] { "Finding", longValue }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the per-row break policy to allow the row-column table row to split.");
        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_DisallowRowBreakRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "Type", "Description" },
                                        new[] { "Finding", longValue }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RowBreakPolicyRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = true;
        style.RowAllowBreakAcrossPages = new List<bool?> { null, false };
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "Type", "Description" },
                                        new[] { "Finding", longValue }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    private static PdfOptions CreateRowSplitRemainderOptions() =>
        new() {
            PageWidth = 360,
            PageHeight = 210,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

    private static PdfTableStyle CreateRowSplitRemainderStyle() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.ColumnWidthPoints = new List<double?> { 70, null };
        return style;
    }

    private static List<string[]> CreateRowSplitRemainderRows() {
        var rows = new List<string[]> {
            new[] { "Type", "Description" }
        };

        for (int index = 1; index <= 6; index++) {
            rows.Add(new[] { "Filler", "Filler row " + index.ToString("00", CultureInfo.InvariantCulture) });
        }

        string splitText = string.Join(" ", Enumerable.Range(1, 34).Select(index =>
            index == 1
                ? "SplitStart01"
                : index == 34
                    ? "SplitTail34"
                    : "SplitMid" + index.ToString("00", CultureInfo.InvariantCulture)));
        rows.Add(new[] { "Finding", splitText });
        rows.Add(new[] { "After", "AfterSplitRow" });
        return rows;
    }

    private static double GetWordTop(PdfPigDocument pdf, int pageNumber, string text) {
        var word = pdf.GetPage(pageNumber)
            .GetWords()
            .FirstOrDefault(candidate => string.Equals(candidate.Text, text, StringComparison.Ordinal));
        Assert.NotNull(word);
        return word!.BoundingBox.Top;
    }

    private static void AssertRowSplitUsesPageRemainder(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the allowed multi-line row to continue onto another page.");

        string firstPage = pdf.GetPage(1).Text;
        string remainingPages = string.Join(Environment.NewLine, Enumerable.Range(2, pdf.NumberOfPages - 1).Select(page => pdf.GetPage(page).Text));

        Assert.Contains("SplitStart01", firstPage);
        Assert.DoesNotContain("SplitTail34", firstPage);
        Assert.Contains("SplitTail34", remainingPages);
        Assert.Contains("AfterSplitRow", remainingPages);
        Assert.Contains("Type", pdf.GetPage(2).Text);
        Assert.Contains("Description", pdf.GetPage(2).Text);
    }


}
