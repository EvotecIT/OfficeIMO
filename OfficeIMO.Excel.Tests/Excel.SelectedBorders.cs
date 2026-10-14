using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Selected_borders_preserve_other_edges_and_font_fill_and_format_after_save() {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Borders");
        sheet.CellAt(1, 1).SetValue(42).SetBold().SetFillColor("ABCDEF");
        sheet.FormatCell(1, 1, "0.00");
        sheet.CellBorder(1, 1, ExcelBorderStyle.Thin, "123456");
        sheet.CellAt(1, 1).SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Thick, "FF0000");
        sheet.CellBorder(1, 1, ExcelCellBorders.Right, ExcelBorderStyle.None);
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.CellBorder(4, 4, (ExcelCellBorders)128, ExcelBorderStyle.Thin));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.CellBorder(4, 4, ExcelCellBorders.Top, (ExcelBorderStyle)999));
        Assert.Throws<ArgumentException>(() => sheet.CellBorder(4, 4, ExcelCellBorders.Top, ExcelBorderStyle.Thin, "FF000O"));
        Assert.Single(sheet.EnumerateCells());
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved); var cell = reopened.Sheets[0].CellAt(1, 1); var style = cell.GetStyle();
        Assert.Equal(42d, cell.GetValue<double>()); Assert.True(style.Bold);
        Assert.Equal("FFABCDEF", style.FillColorArgb); Assert.Equal("0.00", style.NumberFormatCode);
        Assert.Equal("thin", style.Border!.Top!.Style); Assert.Equal("FF123456", style.Border.Top.ColorArgb);
        Assert.Equal("thin", style.Border.Left!.Style); Assert.Equal("FF123456", style.Border.Left.ColorArgb);
        Assert.Equal("thick", style.Border.Bottom!.Style); Assert.Equal("FFFF0000", style.Border.Bottom.ColorArgb);
        Assert.Equal("none", style.Border.Right!.Style); Assert.Null(style.Border.Right.ColorArgb);
    }
}
