using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    internal readonly struct ExcelTableHeaderVisualStyle {
        internal ExcelTableHeaderVisualStyle(string? fillColorArgb, string fontColorArgb, bool bold) {
            FillColorArgb = fillColorArgb;
            FontColorArgb = fontColorArgb;
            Bold = bold;
        }

        internal string? FillColorArgb { get; }

        internal string FontColorArgb { get; }

        internal bool Bold { get; }
    }

    /// <summary>
    /// Projects built-in table header appearance into dependency-free image snapshots.
    /// Excel stores built-in table styles as table metadata rather than cell style indexes,
    /// so managed SVG/PNG rendering needs this narrow visual projection.
    /// </summary>
    internal static class ExcelTableHeaderVisualStyleResolver {
        internal static IReadOnlyDictionary<string, ExcelTableHeaderVisualStyle> Build(
            ExcelSheet sheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn) {
            var styles = new Dictionary<string, ExcelTableHeaderVisualStyle>(StringComparer.OrdinalIgnoreCase);
            foreach (TableDefinitionPart tablePart in sheet.WorksheetPart.TableDefinitionParts) {
                DocumentFormat.OpenXml.Spreadsheet.Table? table = tablePart.Table;
                string? reference = table?.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)
                    || table!.HeaderRowCount?.Value == 0U
                    || !A1.TryParseRange(reference!, out int tableFirstRow, out int tableFirstColumn, out _, out int tableLastColumn)
                    || tableFirstRow < firstRow
                    || tableFirstRow > lastRow) {
                    continue;
                }

                if (!ExcelBuiltInTableStylePaletteResolver.TryCreate(
                        sheet.Document,
                        table.TableStyleInfo?.Name?.Value,
                        out ExcelBuiltInTableStylePalette? palette)) {
                    continue;
                }

                var visualStyle = new ExcelTableHeaderVisualStyle(
                    palette!.HeaderFill,
                    palette.HeaderText ?? "000000",
                    palette.HeaderBold);

                int startColumn = Math.Max(firstColumn, tableFirstColumn);
                int endColumn = Math.Min(lastColumn, tableLastColumn);
                for (int column = startColumn; column <= endColumn; column++) {
                    styles[A1.CellReference(tableFirstRow, column)] = visualStyle;
                }
            }

            return styles;
        }

        internal static ExcelCellStyleSnapshot Apply(ExcelCellStyleSnapshot style, ExcelTableHeaderVisualStyle tableStyle) {
            bool hasDirectStyle = style.StyleIndex != 0U;
            bool hasDirectFontStyle = style.IsFontFamilyExplicit;
            bool preserveDirectFill = hasDirectStyle && !string.IsNullOrWhiteSpace(style.FillColorArgb);
            string? fillColorArgb = preserveDirectFill
                ? style.FillColorArgb
                : tableStyle.FillColorArgb ?? style.FillColorArgb;
            string? fontColorArgb = hasDirectFontStyle
                ? style.FontColorArgb ?? "000000"
                : tableStyle.FontColorArgb;

            return new ExcelCellStyleSnapshot {
                StyleIndex = style.StyleIndex,
                NumberFormatId = style.NumberFormatId,
                NumberFormatCode = style.NumberFormatCode,
                IsDateLike = style.IsDateLike,
                Bold = hasDirectFontStyle ? style.Bold : tableStyle.Bold,
                Italic = style.Italic,
                Underline = style.Underline,
                Strikethrough = style.Strikethrough,
                FontName = style.FontName,
                IsFontFamilyExplicit = style.IsFontFamilyExplicit,
                FontSize = style.FontSize,
                TextRotation = style.TextRotation,
                FontColorArgb = fontColorArgb,
                FillColorArgb = fillColorArgb,
                FillPatternType = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillPatternType : "solid",
                FillPatternForegroundColorArgb = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillPatternForegroundColorArgb : tableStyle.FillColorArgb,
                FillPatternBackgroundColorArgb = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillPatternBackgroundColorArgb : tableStyle.FillColorArgb,
                FillGradientUnsupported = (tableStyle.FillColorArgb == null || preserveDirectFill) && style.FillGradientUnsupported,
                FillGradientStartColorArgb = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillGradientStartColorArgb : null,
                FillGradientEndColorArgb = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillGradientEndColorArgb : null,
                FillGradientStops = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillGradientStops : Array.Empty<ExcelGradientFillStopSnapshot>(),
                FillGradientDegree = tableStyle.FillColorArgb == null || preserveDirectFill ? style.FillGradientDegree : null,
                Border = style.Border,
                HorizontalAlignment = style.HorizontalAlignment,
                VerticalAlignment = style.VerticalAlignment,
                TextIndent = style.TextIndent,
                WrapText = style.WrapText,
                ShrinkToFit = style.ShrinkToFit
            };
        }

    }
}
