using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.ContentSafety;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelCustomIndexedColorTests {
    [Fact]
    public void Workbook_palette_colors_reach_cell_inspection_visual_snapshots_and_saved_readback() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Palette");
        sheet.Cell(1, 1).SetValue("Custom palette");
        sheet.Cell(1, 1).SetFontColor("112233").SetFillColor("445566")
            .SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Thin, "778899");
        var workbook = document.OpenXmlDocument.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet!;
        var cell = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
        var format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value).Color = new Color { Indexed = 0U };
        styles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value).PatternFill!.ForegroundColor = new ForegroundColor { Indexed = 1U };
        var bottom = styles.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value).BottomBorder!;
        bottom.RemoveAllChildren<Color>(); bottom.Append(new Color { Indexed = 2U });
        styles.Colors = new Colors(new IndexedColors(new RgbColor { Rgb = "FF112233" },
            new RgbColor { Rgb = "FF445566" }, new RgbColor { Rgb = "FF778899" }));

        AssertPaletteStyle(sheet.Cell(1, 1).GetStyle());
        AssertPaletteStyle(Assert.Single(Assert.Single(document.Inspect().Worksheets).Cells).Style!);
        AssertPaletteStyle(Assert.Single(sheet.Range("A1").CreateVisualSnapshot().Cells).Style);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        AssertPaletteStyle(reopened.Sheets[0].Cell(1, 1).GetStyle());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void Custom_palette_resolution_applies_tint_and_observes_palette_edits_without_changing_default_colors() {
        using var document = ExcelDocument.Create();
        document.AddSheet("Palette").Cell(1, 1).SetFontColor("112233");
        var workbook = document.OpenXmlDocument.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet!;
        var entry = new RgbColor { Rgb = "FF112233" };
        styles.Colors = new Colors(new IndexedColors(entry));
        Assert.Equal("FF112233", ExcelThemeColorResolver.Resolve(new Color { Indexed = 0U }, workbook));
        Assert.Equal("FFFFFFFF", ExcelThemeColorResolver.Resolve(new Color { Indexed = 0U, Tint = 1d }, workbook));
        entry.Rgb = "FFABCDEF";
        Assert.Equal("FFABCDEF", ExcelThemeColorResolver.Resolve(new Color { Indexed = 0U }, workbook));
        Assert.Null(ExcelThemeColorResolver.Resolve(new Color { Indexed = uint.MaxValue }, workbook));
        styles.Colors = null;
        Assert.Equal("FF000000", ExcelThemeColorResolver.Resolve(new Color { Indexed = 0U }, workbook));
    }

    [Fact]
    public void Content_safety_inspection_uses_the_workbook_palette_for_indexed_font_colors() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Palette");
        sheet.Cell(1, 1).SetValue("Concealed text");
        sheet.Cell(1, 1).SetFontColor("112233");
        var workbook = document.OpenXmlDocument.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet!;
        Cell cell = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
        CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value).Color = new Color { Indexed = 2U };
        styles.Colors = new Colors(new IndexedColors(new RgbColor { Rgb = "FF000000" },
            new RgbColor { Rgb = "FF000000" }, new RgbColor { Rgb = "FFFFFFFF" }));

        byte[] bytes = document.ToBytes(ExcelFileFormat.Xlsx);
        Assert.Contains(ExcelDocument.InspectContentSafety(bytes, "palette.xlsx").Findings,
            finding => finding.Kind == OfficeContentConcealmentKind.LowContrastText
                && finding.Location.IndexOf("Cell(A1)", StringComparison.Ordinal) >= 0);
    }

    private static void AssertPaletteStyle(ExcelCellStyleInfo style) {
        Assert.Equal("FF112233", style.FontColorArgb);
        Assert.Equal("FF445566", style.FillColorArgb);
        Assert.Equal("FF778899", style.Border!.Bottom!.ColorArgb);
    }
}
