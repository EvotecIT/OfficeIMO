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
    public void TableBodyText_ResetsFillColorAfterColoredHeader() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Signal", "Evidence" },
                new[] { "DMARC", "Body text must stay readable after a white header." }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(32, 76, 120),
                HeaderTextColor = PdfColor.White,
                TextColor = null,
                RowStripeFill = PdfColor.FromRgb(248, 250, 252)
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int headerTextColorIndex = content.IndexOf("1 1 1 rg", StringComparison.Ordinal);
        int bodyTextIndex = content.IndexOf("<444D415243> Tj", StringComparison.Ordinal);
        int bodyTextColorIndex = content.LastIndexOf("0 0 0 rg", bodyTextIndex, StringComparison.Ordinal);

        Assert.True(headerTextColorIndex >= 0, "The header should use the configured white text color.");
        Assert.True(bodyTextIndex > headerTextColorIndex, "The body cell should be written after the colored header.");
        Assert.True(bodyTextColorIndex > headerTextColorIndex, "Body cells without an explicit color should reset fill to black.");
    }

    [Fact]
    public void TableStyle_CanControlGenericHeaderBodyAndFooterTypography() {
        var style = TableStyles.Minimal();
        style.FontSize = 8;
        style.LineHeight = 1.1;
        style.HeaderFontSize = 12;
        style.FooterFontSize = 10;
        style.FooterRowCount = 1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Column", "Status" },
                new[] { "BodyRow", "Readable" },
                new[] { "Total", "Footer" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/F2 12 Tf", content);
        Assert.Contains("/F1 8 Tf", content);
        Assert.Contains("/F2 10 Tf", content);
    }

    [Fact]
    public void TableStyle_UsesConfiguredCellLineHeight() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTableLineHeightProbe(options, null, useRowColumnFlow: false);
        byte[] looseBytes = CreateTableLineHeightProbe(options, 2.0, useRowColumnFlow: false);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfPigDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultGap = FindWordStartY(defaultPage, "FirstLine") - FindWordStartY(defaultPage, "SecondLine");
        double looseGap = FindWordStartY(loosePage, "FirstLine") - FindWordStartY(loosePage, "SecondLine");

        Assert.True(looseGap > defaultGap + 4, $"Expected larger table line height to increase wrapped cell baseline gap. Default: {defaultGap:0.##}, loose: {looseGap:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredCellLineHeight() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTableLineHeightProbe(options, null, useRowColumnFlow: true);
        byte[] looseBytes = CreateTableLineHeightProbe(options, 2.0, useRowColumnFlow: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfPigDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultGap = FindWordStartY(defaultPage, "FirstLine") - FindWordStartY(defaultPage, "SecondLine");
        double looseGap = FindWordStartY(loosePage, "FirstLine") - FindWordStartY(loosePage, "SecondLine");

        Assert.True(looseGap > defaultGap + 4, $"Expected larger row-column table line height to increase wrapped cell baseline gap. Default: {defaultGap:0.##}, loose: {looseGap:0.##}.");
    }

    [Fact]
    public void TableStyle_UsesConfiguredCellPaddingSides() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePaddingProbe(options, useRowColumnFlow: false, useSidePadding: false);
        byte[] paddedBytes = CreateTablePaddingProbe(options, useRowColumnFlow: false, useSidePadding: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfPigDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "PadMarker");
        double paddedX = FindWordStartX(paddedPage, "PadMarker");
        double defaultY = FindWordStartY(defaultPage, "PadMarker");
        double paddedY = FindWordStartY(paddedPage, "PadMarker");

        Assert.True(paddedX > defaultX + 14, $"Expected left cell padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected top cell padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredCellPaddingSides() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePaddingProbe(options, useRowColumnFlow: true, useSidePadding: false);
        byte[] paddedBytes = CreateTablePaddingProbe(options, useRowColumnFlow: true, useSidePadding: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfPigDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "PadMarker");
        double paddedX = FindWordStartX(paddedPage, "PadMarker");
        double defaultY = FindWordStartY(defaultPage, "PadMarker");
        double paddedY = FindWordStartY(paddedPage, "PadMarker");

        Assert.True(paddedX > defaultX + 14, $"Expected row-column left cell padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected row-column top cell padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void TableStyle_UsesConfiguredPerCellPadding() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePerCellPaddingProbe(options, useRowColumnFlow: false, useCellPadding: false);
        byte[] paddedBytes = CreateTablePerCellPaddingProbe(options, useRowColumnFlow: false, useCellPadding: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfPigDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "CellPadMarker");
        double paddedX = FindWordStartX(paddedPage, "CellPadMarker");
        double defaultY = FindWordStartY(defaultPage, "CellPadMarker");
        double paddedY = FindWordStartY(paddedPage, "CellPadMarker");

        Assert.True(paddedX > defaultX + 16, $"Expected per-cell left padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected per-cell top padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredPerCellPadding() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePerCellPaddingProbe(options, useRowColumnFlow: true, useCellPadding: false);
        byte[] paddedBytes = CreateTablePerCellPaddingProbe(options, useRowColumnFlow: true, useCellPadding: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfPigDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "CellPadMarker");
        double paddedX = FindWordStartX(paddedPage, "CellPadMarker");
        double defaultY = FindWordStartY(defaultPage, "CellPadMarker");
        double paddedY = FindWordStartY(paddedPage, "CellPadMarker");

        Assert.True(paddedX > defaultX + 16, $"Expected row-column per-cell left padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected row-column per-cell top padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void TableStyle_UsesConfiguredPerCellAlignment() {
        var options = new PdfOptions {
            PageWidth = 280,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePerCellAlignmentProbe(options, useRowColumnFlow: false, useCellAlignment: false);
        byte[] alignedBytes = CreateTablePerCellAlignmentProbe(options, useRowColumnFlow: false, useCellAlignment: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var alignedPdf = PdfPigDocument.Open(new MemoryStream(alignedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var alignedPage = alignedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "CellAlignMarker");
        double alignedX = FindWordStartX(alignedPage, "CellAlignMarker");
        double defaultY = FindWordStartY(defaultPage, "CellAlignMarker");
        double alignedY = FindWordStartY(alignedPage, "CellAlignMarker");

        Assert.True(alignedX > defaultX + 30, $"Expected per-cell right alignment to move text right. Default x: {defaultX:0.##}, aligned x: {alignedX:0.##}.");
        Assert.True(defaultY > alignedY + 30, $"Expected per-cell bottom alignment to move text down. Default y: {defaultY:0.##}, aligned y: {alignedY:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredPerCellAlignment() {
        var options = new PdfOptions {
            PageWidth = 280,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePerCellAlignmentProbe(options, useRowColumnFlow: true, useCellAlignment: false);
        byte[] alignedBytes = CreateTablePerCellAlignmentProbe(options, useRowColumnFlow: true, useCellAlignment: true);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var alignedPdf = PdfPigDocument.Open(new MemoryStream(alignedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var alignedPage = alignedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "CellAlignMarker");
        double alignedX = FindWordStartX(alignedPage, "CellAlignMarker");
        double defaultY = FindWordStartY(defaultPage, "CellAlignMarker");
        double alignedY = FindWordStartY(alignedPage, "CellAlignMarker");

        Assert.True(alignedX > defaultX + 30, $"Expected row-column per-cell right alignment to move text right. Default x: {defaultX:0.##}, aligned x: {alignedX:0.##}.");
        Assert.True(defaultY > alignedY + 30, $"Expected row-column per-cell bottom alignment to move text down. Default y: {defaultY:0.##}, aligned y: {alignedY:0.##}.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableStyle_UsesConfiguredCellSpacing(bool useRowColumnFlow) {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTableCellSpacingProbe(options, 0, useRowColumnFlow);
        byte[] spacedBytes = CreateTableCellSpacingProbe(options, 12, useRowColumnFlow);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultHorizontalGap = FindWordStartX(defaultPage, "SpacingB1") - FindWordStartX(defaultPage, "SpacingA1");
        double spacedHorizontalGap = FindWordStartX(spacedPage, "SpacingB1") - FindWordStartX(spacedPage, "SpacingA1");
        double defaultVerticalGap = FindWordStartY(defaultPage, "SpacingA1") - FindWordStartY(defaultPage, "SpacingA2");
        double spacedVerticalGap = FindWordStartY(spacedPage, "SpacingA1") - FindWordStartY(spacedPage, "SpacingA2");

        Assert.True(spacedHorizontalGap > defaultHorizontalGap + 10, $"Expected cell spacing to increase horizontal cell distance. Default: {defaultHorizontalGap:0.##}, spaced: {spacedHorizontalGap:0.##}.");
        Assert.True(spacedVerticalGap > defaultVerticalGap + 10, $"Expected cell spacing to increase vertical row distance. Default: {defaultVerticalGap:0.##}, spaced: {spacedVerticalGap:0.##}.");
    }

    [Fact]
    public void TableStyle_CanDisableHeaderAndFooterBoldWithoutChangingDocumentFont() {
        var style = TableStyles.Minimal();
        style.HeaderBold = false;
        style.FooterBold = false;
        style.FooterRowCount = 1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "PlainHeader", "Status" },
                new[] { "BodyRow", "Readable" },
                new[] { "PlainFooter", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/F1 9 Tf", content);
        Assert.DoesNotContain("/F2 9 Tf", content);
    }
}
