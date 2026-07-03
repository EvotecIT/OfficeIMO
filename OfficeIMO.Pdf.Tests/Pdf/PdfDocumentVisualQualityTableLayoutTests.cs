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
    public void Table_UsesConfiguredMinimumRowHeight() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.MinRowHeight = 36;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" },
                new[] { "Gamma", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaY = FindWordStartY(page, "Alpha");
        double betaY = FindWordStartY(page, "Beta");
        double gammaY = FindWordStartY(page, "Gamma");

        Assert.True(alphaY - betaY >= 34, $"Expected minimum row height spacing between first and second row. Alpha y: {alphaY:0.##}, Beta y: {betaY:0.##}.");
        Assert.True(betaY - gammaY >= 34, $"Expected minimum row height spacing between second and third row. Beta y: {betaY:0.##}, Gamma y: {gammaY:0.##}.");
    }

    [Fact]
    public void Table_UsesConfiguredPerRowMinimumHeights() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.MinRowHeight = 18;
        style.RowMinHeights = new List<double?> { 18, 54, 18 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" },
                new[] { "Gamma", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaY = FindWordStartY(page, "Alpha");
        double betaY = FindWordStartY(page, "Beta");
        double gammaY = FindWordStartY(page, "Gamma");

        Assert.InRange(alphaY - betaY, 16D, 28D);
        Assert.True(betaY - gammaY >= 52D, $"Expected second row-specific minimum height to push the third row down. Beta y: {betaY:0.##}, Gamma y: {gammaY:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesConfiguredPerRowMinimumHeights() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.MinRowHeight = 18;
        style.RowMinHeights = new List<double?> { 18, 54, 18 };
        var rows = new[] {
            new[] { "Alpha", "Ready" },
            new[] { "Beta", "Ready" },
            new[] { "Gamma", "Ready" }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(compose =>
                compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaY = FindWordStartY(page, "Alpha");
        double betaY = FindWordStartY(page, "Beta");
        double gammaY = FindWordStartY(page, "Gamma");

        Assert.InRange(alphaY - betaY, 16D, 28D);
        Assert.True(betaY - gammaY >= 52D, $"Expected row-column table row-specific minimum height to push the third row down. Beta y: {betaY:0.##}, Gamma y: {gammaY:0.##}.");
    }

    [Fact]
    public void Table_UsesConfiguredSpacingBeforeAndAfter() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        byte[] defaultBytes = CreateTableSpacingProbe(options, spacingBefore: 0, spacingAfter: 0);
        byte[] spacedBytes = CreateTableSpacingProbe(options, spacingBefore: 12, spacingAfter: 18);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultTableY = FindWordStartY(defaultPage, "Alpha");
        double spacedTableY = FindWordStartY(spacedPage, "Alpha");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultTableY - spacedTableY >= 10, $"Expected table spacing before to move table content down. Default y: {defaultTableY:0.##}, spaced y: {spacedTableY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 28, $"Expected table spacing before and after to move following content down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void TableStylesLight_ProvidesDefaultFlowRhythmAroundTables() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9,
            DefaultParagraphStyle = new PdfParagraphStyle {
                SpacingAfter = 0
            }
        };
        PdfTableStyle defaultLight = TableStyles.Light();
        var cramped = TableStyles.Light();
        cramped.SpacingBefore = 0;
        cramped.SpacingAfter = 0;

        byte[] defaultBytes = CreateLightTableRhythmProbe(options, style: null);
        byte[] crampedBytes = CreateLightTableRhythmProbe(options, cramped);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var crampedPdf = PdfPigDocument.Open(new MemoryStream(crampedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var crampedPage = crampedPdf.GetPage(1);

        double defaultTableY = FindWordStartY(defaultPage, "Alpha");
        double crampedTableY = FindWordStartY(crampedPage, "Alpha");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double crampedAfterY = FindWordStartY(crampedPage, "AfterMarker");

        Assert.Equal(4, defaultLight.SpacingBefore);
        Assert.Equal(8, defaultLight.SpacingAfter);
        Assert.True(crampedTableY - defaultTableY >= 3, $"Expected default light-table spacing before to move table content down. Cramped y: {crampedTableY:0.##}, default y: {defaultTableY:0.##}.");
        Assert.True(crampedAfterY - defaultAfterY >= 10, $"Expected default light-table rhythm to separate following paragraphs from the grid. Cramped y: {crampedAfterY:0.##}, default y: {defaultAfterY:0.##}.");
    }

    [Fact]
    public void Table_SuppressesSpacingBeforeAtPageTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.HeaderRowCount = 0;
        var spacedStyle = TableStyles.Minimal();
        spacedStyle.HeaderRowCount = 0;
        spacedStyle.SpacingBefore = 28;
        spacedStyle.SpacingAfter = 0;

        byte[] defaultBytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "TopTableMarker", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "TopTableMarker", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopTableMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopTableMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void RowColumnTable_SuppressesSpacingBeforeAtColumnTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.HeaderRowCount = 0;
        var spacedStyle = TableStyles.Minimal();
        spacedStyle.HeaderRowCount = 0;
        spacedStyle.SpacingBefore = 28;
        spacedStyle.SpacingAfter = 0;
        string[][] rows = {
            new[] { "ColumnTableMarker", "Ready" },
            new[] { "Beta", "Ready" }
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Table(rows, style: defaultStyle))))))
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Table(rows, style: spacedStyle))))))
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "ColumnTableMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "ColumnTableMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Table_RendersConfiguredCaptionAboveGrid() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.Caption = "SignalCaption";
        style.CaptionAlign = PdfAlign.Right;
        style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = 8;
        style.CaptionSpacingAfter = 10;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double captionY = FindWordStartY(page, "SignalCaption");
        double alphaY = FindWordStartY(page, "Alpha");
        double captionX = FindWordStartX(page, "SignalCaption");
        double alphaX = FindWordStartX(page, "Alpha");

        Assert.True(captionY > alphaY + 14, $"Expected the table caption above the first row. Caption y: {captionY:0.##}, Alpha y: {alphaY:0.##}.");
        Assert.True(captionX > alphaX + 120, $"Expected the right-aligned caption to render near the table's right edge. Caption x: {captionX:0.##}, Alpha x: {alphaX:0.##}.");

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.314 0.353 0.392 rg", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCaptionAboveGrid() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.Caption = "SignalCaption";
        style.CaptionAlign = PdfAlign.Right;
        style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = 8;
        style.CaptionSpacingAfter = 10;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Ready" },
                                    new[] { "Beta", "Ready" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double captionY = FindWordStartY(page, "SignalCaption");
        double alphaY = FindWordStartY(page, "Alpha");
        double captionX = FindWordStartX(page, "SignalCaption");
        double alphaX = FindWordStartX(page, "Alpha");

        Assert.True(captionY > alphaY + 14, $"Expected the row-column table caption above the first row. Caption y: {captionY:0.##}, Alpha y: {alphaY:0.##}.");
        Assert.True(captionX > alphaX + 120, $"Expected the right-aligned row-column caption to render near the table's right edge. Caption x: {captionX:0.##}, Alpha x: {alphaX:0.##}.");

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.314 0.353 0.392 rg", content);
    }

    [Fact]
    public void Table_RejectsCaptionAndFirstRowTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.Caption = string.Join(" ", Enumerable.Repeat("TallCaption", 40));
        style.CaptionFontSize = 14;
        style.CaptionSpacingAfter = 8;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Table(new[] {
                    new[] { "Alpha", "Ready" },
                    new[] { "Beta", "Ready" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table caption and first row exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsCaptionAndFirstRowTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.Caption = string.Join(" ", Enumerable.Repeat("TallCaption", 40));
        style.CaptionFontSize = 14;
        style.CaptionSpacingAfter = 8;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "Alpha", "Ready" },
                                        new[] { "Beta", "Ready" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table caption and first row exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_UsesRelativeColumnWidthWeights() {
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
        style.ColumnWidthWeights = new List<double> { 1, 3, 1 };

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
        Assert.True(secondColumnWidth > firstColumnWidth * 2, $"Expected the middle table column to be visibly wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesRelativeColumnWidthWeights() {
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
        style.ColumnWidthWeights = new List<double> { 1, 3, 1 };

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
        Assert.True(secondColumnWidth > firstColumnWidth * 2, $"Expected the row-column middle table column to be visibly wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void Table_MaxWidthCapsWeightedColumnsAndHonorsAlignment() {
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
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Beta" }
            }, align: PdfAlign.Right, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 152, 158);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void RowColumnTable_MaxWidthCapsWeightedColumnsAndHonorsAlignment() {
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
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Beta" }
                                }, align: PdfAlign.Center, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 92, 98);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void Table_LeftIndentOffsetsTableFrameBeforeColumnSizing() {
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
        style.LeftIndent = 60;
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Beta" }
            }, align: PdfAlign.Left, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 92, 98);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void RowColumnTable_LeftIndentOffsetsTableFrameBeforeColumnSizing() {
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
        style.LeftIndent = 40;
        style.MaxWidth = 120;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Beta" }
                                }, align: PdfAlign.Left, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 72, 78);
        Assert.InRange(betaX - alphaX, 38, 48);
    }


}
