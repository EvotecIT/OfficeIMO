using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Applies a border style to selected edges of a one-based cell, preserving other edges and style facets.
    /// <see cref="ExcelBorderStyle.None"/> explicitly clears the selected edges.</summary>
    /// <param name="row">One-based row.</param>
    /// <param name="column">One-based column.</param>
    /// <param name="borders">The rectangular edges to change.</param>
    /// <param name="style">The border style.</param>
    /// <param name="hexColor">Optional RGB or ARGB hex color.</param>
    public void CellBorder(int row, int column, ExcelCellBorders borders, ExcelBorderStyle style, string? hexColor = null) {
        if ((borders & ~ExcelCellBorders.All) != 0) throw new ArgumentOutOfRangeException(nameof(borders));
        if (borders == ExcelCellBorders.None) return;
        BorderStyleValues value = style.ToOpenXml();
        string? argb = NormalizeOptionalHexColor(hexColor);
        WriteLockConditional(() => {
            var cell = GetCell(row, column);
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            var baseFormat = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
            uint borderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(baseFormat.BorderId), border => {
                if (borders.HasFlag(ExcelCellBorders.Left)) border.LeftBorder = CreateBorderSide<LeftBorder>(value, argb);
                if (borders.HasFlag(ExcelCellBorders.Top)) border.TopBorder = CreateBorderSide<TopBorder>(value, argb);
                if (borders.HasFlag(ExcelCellBorders.Right)) border.RightBorder = CreateBorderSide<RightBorder>(value, argb);
                if (borders.HasFlag(ExcelCellBorders.Bottom)) border.BottomBorder = CreateBorderSide<BottomBorder>(value, argb);
            });
            ApplyCellFormatOverride(stylesheet, cell, format => { format.BorderId = borderId; format.ApplyBorder = true; });
            SaveStylesheet(stylesPart);
        });
    }
}
