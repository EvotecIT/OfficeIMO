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
    public void TableStyles_ExposeWordLikeGenericPresetsWithoutSemanticAlignment() {
        var tableGrid = TableStyles.TableGrid();
        var tableGridLight = TableStyles.TableGridLight();
        var plainTable = TableStyles.PlainTable1();
        var gridTable = TableStyles.GridTable1Light();
        var listTable = TableStyles.ListTable1Light();

        Assert.Equal(PdfColor.Black, tableGrid.BorderColor);
        Assert.Equal(0.5, tableGrid.BorderWidth);
        Assert.Null(tableGrid.HeaderFill);
        Assert.Null(tableGrid.RowStripeFill);

        Assert.Equal(PdfColor.FromRgb(191, 191, 191), tableGridLight.BorderColor);
        Assert.Equal(0.5, tableGridLight.BorderWidth);
        Assert.Null(tableGridLight.HeaderFill);
        Assert.Null(tableGridLight.RowStripeFill);
        Assert.NotEqual(tableGrid.BorderColor, tableGridLight.BorderColor);

        Assert.Null(plainTable.BorderColor);
        Assert.Equal(0, plainTable.BorderWidth);
        Assert.Null(plainTable.RowSeparatorColor);
        Assert.Null(plainTable.HeaderSeparatorColor);

        Assert.Equal(PdfColor.FromRgb(217, 217, 217), gridTable.BorderColor);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridTable.HeaderSeparatorColor);
        Assert.Equal(0.8, gridTable.HeaderSeparatorWidth);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridTable.FooterSeparatorColor);
        Assert.Equal(0.8, gridTable.FooterSeparatorWidth);

        Assert.Null(listTable.BorderColor);
        Assert.Equal(PdfColor.Black, listTable.HeaderSeparatorColor);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(0.8, listTable.FooterSeparatorWidth);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTable.RowSeparatorColor);

        Assert.False(tableGrid.RightAlignNumeric);
        Assert.False(tableGridLight.RightAlignNumeric);
        Assert.False(plainTable.RightAlignNumeric);
        Assert.False(gridTable.RightAlignNumeric);
        Assert.False(listTable.RightAlignNumeric);

        var independentGridTable = TableStyles.GridTable1Light();
        gridTable.CellPaddingX = 20;
        Assert.Equal(5, independentGridTable.CellPaddingX);
    }

    [Fact]
    public void TableStyles_ResolveSupportedWordStyleNamesToFreshPdfStyles() {
        Assert.Equal(new[] {
            "TableNormal",
            "TableGrid",
            "TableGridLight",
            "PlainTable1",
            "GridTable1Light",
            "GridTable1LightAccent1",
            "GridTable1LightAccent2",
            "GridTable1LightAccent3",
            "GridTable1LightAccent4",
            "GridTable1LightAccent5",
            "GridTable1LightAccent6",
            "ListTable1Light",
            "ListTable1LightAccent1",
            "ListTable1LightAccent2",
            "ListTable1LightAccent3",
            "ListTable1LightAccent4",
            "ListTable1LightAccent5",
            "ListTable1LightAccent6",
            "GridTableLight",
            "GridTable1Light-Accent1",
            "GridTable1Light-Accent2",
            "GridTable1Light-Accent3",
            "GridTable1Light-Accent4",
            "GridTable1Light-Accent5",
            "GridTable1Light-Accent6",
            "ListTable1Light-Accent1",
            "ListTable1Light-Accent2",
            "ListTable1Light-Accent3",
            "ListTable1Light-Accent4",
            "ListTable1Light-Accent5",
            "ListTable1Light-Accent6"
        }, TableStyles.SupportedWordStyleNames);

        PdfTableStyle tableNormal = TableStyles.FromWordTableStyle("Table Normal");
        PdfTableStyle tableGrid = TableStyles.FromWordTableStyle("Table Grid");
        PdfTableStyle tableGridLight = TableStyles.FromWordTableStyle("Grid Table Light");
        PdfTableStyle plainTable = TableStyles.FromWordTableStyle("plain_table_1");
        bool resolvedGridLight = TableStyles.TryFromWordTableStyle("grid-table-1-light", out PdfTableStyle? gridLight);
        PdfTableStyle gridLightAccent = TableStyles.FromWordTableStyle("GridTable1Light-Accent2");
        PdfTableStyle listTable = TableStyles.FromWordTableStyle(" list table 1 light ");
        PdfTableStyle listTableAccent = TableStyles.FromWordTableStyle("ListTable1LightAccent5");

        Assert.Null(tableNormal.BorderColor);
        Assert.Equal(PdfColor.Black, tableGrid.BorderColor);
        Assert.Equal(PdfColor.FromRgb(191, 191, 191), tableGridLight.BorderColor);
        Assert.Null(tableGridLight.HeaderSeparatorColor);
        Assert.Null(plainTable.BorderColor);
        Assert.True(resolvedGridLight);
        Assert.NotNull(gridLight);
        Assert.Equal(PdfColor.FromRgb(217, 217, 217), gridLight!.BorderColor);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridLight.FooterSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(247, 202, 172), gridLightAccent.BorderColor);
        Assert.Equal(PdfColor.FromRgb(244, 176, 131), gridLightAccent.HeaderSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTable.RowSeparatorColor);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(222, 234, 246), listTableAccent.RowStripeFill);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTableAccent.RowSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(156, 194, 229), listTableAccent.HeaderSeparatorColor);

        PdfTableStyle independentListTable = TableStyles.FromWordTableStyle("ListTable1Light");
        listTable.CellPaddingX = 20;
        Assert.Equal(4, independentListTable.CellPaddingX);

        Assert.False(TableStyles.TryFromWordTableStyle("GridTable7Colorful", out PdfTableStyle? missingStyle));
        Assert.Null(missingStyle);

        var exception = Assert.Throws<ArgumentException>(() => TableStyles.FromWordTableStyle("GridTable7Colorful"));
        Assert.Equal("styleName", exception.ParamName);
        Assert.Contains("Unsupported Word table style 'GridTable7Colorful'.", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Supported styles: TableNormal, TableGrid, TableGridLight, PlainTable1, GridTable1Light", exception.Message, StringComparison.Ordinal);
        Assert.Contains("GridTable1Light-Accent6", exception.Message, StringComparison.Ordinal);
        Assert.Contains("ListTable1Light-Accent6", exception.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentNullException>(() => TableStyles.FromWordTableStyle(null!));
        Assert.Throws<ArgumentNullException>(() => TableStyles.TryFromWordTableStyle(null!, out _));
    }

    [Fact]
    public void TableStyles_ExposeCanonicalWordStyleNamesWithoutAliasSpellings() {
        Assert.Contains("TableNormal", TableStyles.CanonicalWordStyleNames);
        Assert.Contains("TableGridLight", TableStyles.CanonicalWordStyleNames);
        Assert.Contains("GridTable1LightAccent6", TableStyles.CanonicalWordStyleNames);
        Assert.Contains("ListTable1LightAccent6", TableStyles.CanonicalWordStyleNames);
        Assert.DoesNotContain("GridTable1Light-Accent1", TableStyles.CanonicalWordStyleNames);
        Assert.DoesNotContain("ListTable1Light-Accent1", TableStyles.CanonicalWordStyleNames);
        Assert.DoesNotContain("GridTableLight", TableStyles.CanonicalWordStyleNames);

        Assert.Contains("GridTableLight", TableStyles.SupportedWordStyleNames);
        Assert.Contains("GridTable1Light-Accent1", TableStyles.SupportedWordStyleNames);
        Assert.Contains("ListTable1Light-Accent1", TableStyles.SupportedWordStyleNames);
        Assert.Equal(TableStyles.CanonicalWordStyleNames.Count, TableStyles.CanonicalWordStyleNames.Distinct(StringComparer.Ordinal).Count());
    }

    [Theory]
    [InlineData("Table Normal", "TableNormal")]
    [InlineData("table-grid", "TableGrid")]
    [InlineData("Grid Table Light", "TableGridLight")]
    [InlineData("table_grid_light", "TableGridLight")]
    [InlineData("plain_table_1", "PlainTable1")]
    [InlineData("grid-table-1-light", "GridTable1Light")]
    [InlineData("GridTable1Light-Accent2", "GridTable1LightAccent2")]
    [InlineData("grid table 1 light accent 6", "GridTable1LightAccent6")]
    [InlineData("ListTable1Light-Accent5", "ListTable1LightAccent5")]
    [InlineData(" list table 1 light accent 3 ", "ListTable1LightAccent3")]
    public void TableStyles_NormalizeSupportedWordStyleNamesToCanonicalNames(string input, string expectedCanonicalName) {
        Assert.True(TableStyles.TryGetCanonicalWordStyleName(input, out string? canonicalName));
        Assert.Equal(expectedCanonicalName, canonicalName);
        Assert.Equal(expectedCanonicalName, TableStyles.GetCanonicalWordStyleName(input));
    }

    [Fact]
    public void TableStyles_CanonicalWordStyleNameRejectsUnsupportedInputs() {
        Assert.False(TableStyles.TryGetCanonicalWordStyleName("GridTable7Colorful", out string? missingStyle));
        Assert.Null(missingStyle);

        var exception = Assert.Throws<ArgumentException>(() => TableStyles.GetCanonicalWordStyleName("GridTable7Colorful"));
        Assert.Equal("styleName", exception.ParamName);
        Assert.Contains("Unsupported Word table style 'GridTable7Colorful'.", exception.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentNullException>(() => TableStyles.GetCanonicalWordStyleName(null!));
        Assert.Throws<ArgumentNullException>(() => TableStyles.TryGetCanonicalWordStyleName(null!, out _));
    }

    [Theory]
    [InlineData(1, 180, 198, 231, 142, 170, 219, 217, 226, 243)]
    [InlineData(2, 247, 202, 172, 244, 176, 131, 251, 228, 213)]
    [InlineData(3, 219, 219, 219, 201, 201, 201, 237, 237, 237)]
    [InlineData(4, 255, 229, 153, 255, 217, 102, 255, 242, 204)]
    [InlineData(5, 189, 214, 238, 156, 194, 229, 222, 234, 246)]
    [InlineData(6, 197, 224, 179, 168, 208, 141, 226, 239, 217)]
    public void TableStyles_ResolveWordAccentVariantsWithDefaultThemeColors(
        int accent,
        int lightR,
        int lightG,
        int lightB,
        int strongR,
        int strongG,
        int strongB,
        int paleR,
        int paleG,
        int paleB) {
        PdfTableStyle grid = TableStyles.FromWordTableStyle("GridTable1Light-Accent" + accent.ToString(CultureInfo.InvariantCulture));
        PdfTableStyle list = TableStyles.FromWordTableStyle("ListTable1LightAccent" + accent.ToString(CultureInfo.InvariantCulture));

        PdfColor ExpectedRgb(int r, int g, int b) => PdfColor.FromRgb((byte)r, (byte)g, (byte)b);

        Assert.Equal(ExpectedRgb(lightR, lightG, lightB), grid.BorderColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), grid.HeaderSeparatorColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), grid.FooterSeparatorColor);

        Assert.Equal(ExpectedRgb(paleR, paleG, paleB), list.RowStripeFill);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), list.HeaderSeparatorColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), list.FooterSeparatorColor);
    }

    [Fact]
    public void TableStyles_WordLikePresetsRenderDistinctGridAndListGeometry() {
        string plainContent = RenderTableStyleContent(TableStyles.PlainTable1());
        string gridContent = RenderTableStyleContent(TableStyles.TableGrid());
        string gridLightContent = RenderTableStyleContent(TableStyles.GridTable1Light());
        string listContent = RenderTableStyleContent(TableStyles.ListTable1Light());

        Assert.DoesNotContain(" re S", plainContent);
        Assert.DoesNotContain(" l S", plainContent);

        Assert.Contains(" re S", gridContent);
        Assert.Contains(" l S", gridContent);

        Assert.Contains(" re S", gridLightContent);
        Assert.Contains(" l S", gridLightContent);
        Assert.Contains("0.8 w", gridLightContent);

        Assert.DoesNotContain(" re S", listContent);
        Assert.Contains(" l S", listContent);
        Assert.Contains("0.45 w", listContent);
        Assert.Contains("0.8 w", listContent);
    }

    [Fact]
    public void WordLikeTablePresets_ProvideFooterSeparatorsForSummaryRows() {
        PdfTableStyle gridLight = TableStyles.GridTable1Light();
        PdfTableStyle listTable = TableStyles.ListTable1Light();

        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridLight.FooterSeparatorColor);
        Assert.Equal(0.8, gridLight.FooterSeparatorWidth);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(0.8, listTable.FooterSeparatorWidth);
    }

    [Fact]
    public void PdfTheme_WordLikeTableStyleIncludesFooterSeparatorByDefault() {
        PdfTheme theme = PdfTheme.WordLike();
        PdfTableStyle style = theme.TableStyle!;

        Assert.Equal(PdfColor.FromRgb(17, 24, 39), style.FooterSeparatorColor);
        Assert.Equal(0.8, style.FooterSeparatorWidth);
    }

    [Fact]
    public void TableStyles_DocumentPresetsExposeReusableVisualRhythm() {
        PdfTableStyle technical = TableStyles.TechnicalDocument();
        PdfTableStyle compact = TableStyles.Compact();
        PdfTableStyle report = TableStyles.Report();
        PdfTableStyle freshReport = TableStyles.Report();

        Assert.Equal(PdfColor.FromRgb(15, 23, 42), technical.HeaderFill);
        Assert.Equal(PdfColor.White, technical.HeaderTextColor);
        Assert.Equal(PdfColor.FromRgb(226, 232, 240), technical.RowSeparatorColor);
        Assert.Equal(9.75, technical.FontSize);
        Assert.Equal(1.2, technical.LineHeight);
        Assert.Equal(6, technical.SpacingBefore);
        Assert.True(technical.AutoFitColumns);

        Assert.Null(compact.HeaderFill);
        Assert.Equal(9, compact.FontSize);
        Assert.Equal(1.12, compact.LineHeight);
        Assert.Equal(4, compact.CellPaddingX);
        Assert.Equal(3, compact.CellPaddingY);
        Assert.Equal(7, compact.SpacingAfter);
        Assert.True(compact.AutoFitColumns);

        Assert.Equal(PdfColor.FromRgb(30, 64, 175), report.HeaderFill);
        Assert.Equal(PdfColor.FromRgb(239, 246, 255), report.RowStripeFill);
        Assert.Equal(PdfColor.FromRgb(191, 219, 254), report.BorderColor);
        Assert.Equal(9.25, report.FontSize);
        Assert.Equal(1.18, report.LineHeight);
        Assert.Equal(10, report.SpacingAfter);
        Assert.True(report.AutoFitColumns);

        report.HeaderFill = PdfColor.Black;
        Assert.Equal(PdfColor.FromRgb(30, 64, 175), freshReport.HeaderFill);
    }


}
