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
    public void Table_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the row-column unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void Table_UsesMinimumColumnWidthPointsForWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void RowColumnTable_UsesMinimumColumnWidthPointsForWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void Table_UsesMaximumColumnWidthPointsForWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void RowColumnTable_UsesMaximumColumnWidthPointsForWeightedColumns() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void Table_UsesConfiguredVerticalColumnAlignment() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Name", "Notes" },
                new[] {
                    "BottomValue",
                    "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first-column value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesConfiguredVerticalColumnAlignment() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Notes" },
                                    new[] {
                                        "BottomValue",
                                        "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                                    }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first row-column cell value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsFlexibleColumnsFromMeasuredContent() {
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
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "SKU", "Description", "Amount" },
                new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void RowColumnTable_AutoFitsFlexibleColumnsFromMeasuredContent() {
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
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "SKU", "Description", "Amount" },
                                    new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the row-column description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the row-column description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void Table_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Metric", "Amount" },
                new[] { "Revenue", "$1,234.50" },
                new[] { "Refund", "(45.20)" },
                new[] { "Margin", "99%" },
                new[] { "EU", "€1,234.50" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }

    [Fact]
    public void RowColumnTable_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Amount" },
                                    new[] { "Revenue", "$1,234.50" },
                                    new[] { "Refund", "(45.20)" },
                                    new[] { "Margin", "99%" },
                                    new[] { "EU", "€1,234.50" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }


}
