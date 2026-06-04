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


}
