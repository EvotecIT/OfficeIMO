using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    internal readonly struct ExcelTableHeaderVisualStyle {
        internal ExcelTableHeaderVisualStyle(string? fillColorArgb, string fontColorArgb, bool bold, string? borderColorArgb) {
            FillColorArgb = fillColorArgb;
            FontColorArgb = fontColorArgb;
            Bold = bold;
            BorderColorArgb = borderColorArgb;
        }

        internal string? FillColorArgb { get; }

        internal string FontColorArgb { get; }

        internal bool Bold { get; }

        internal string? BorderColorArgb { get; }
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
                    palette.HeaderBold,
                    palette.Border);

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
            bool preserveDirectFill = hasDirectStyle && HasDirectFill(style);
            bool useTableFill = !preserveDirectFill && !string.IsNullOrWhiteSpace(tableStyle.FillColorArgb);
            bool preserveDirectBorder = hasDirectStyle && style.Border != null;
            string? fillColorArgb = useTableFill ? tableStyle.FillColorArgb : style.FillColorArgb;
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
                FillPatternType = useTableFill ? "solid" : style.FillPatternType,
                FillPatternForegroundColorArgb = useTableFill ? tableStyle.FillColorArgb : style.FillPatternForegroundColorArgb,
                FillPatternBackgroundColorArgb = useTableFill ? tableStyle.FillColorArgb : style.FillPatternBackgroundColorArgb,
                FillGradientUnsupported = !useTableFill && style.FillGradientUnsupported,
                FillGradientStartColorArgb = useTableFill ? null : style.FillGradientStartColorArgb,
                FillGradientEndColorArgb = useTableFill ? null : style.FillGradientEndColorArgb,
                FillGradientStops = useTableFill ? Array.Empty<ExcelGradientFillStopSnapshot>() : style.FillGradientStops,
                FillGradientDegree = useTableFill ? null : style.FillGradientDegree,
                Border = preserveDirectBorder ? style.Border : CreateTableBorder(tableStyle.BorderColorArgb),
                HorizontalAlignment = style.HorizontalAlignment,
                VerticalAlignment = style.VerticalAlignment,
                TextIndent = style.TextIndent,
                WrapText = style.WrapText,
                ShrinkToFit = style.ShrinkToFit
            };
        }

        private static bool HasDirectFill(ExcelCellStyleSnapshot style) =>
            !string.IsNullOrWhiteSpace(style.FillColorArgb) ||
            !string.IsNullOrWhiteSpace(style.FillPatternType) ||
            !string.IsNullOrWhiteSpace(style.FillPatternForegroundColorArgb) ||
            !string.IsNullOrWhiteSpace(style.FillPatternBackgroundColorArgb) ||
            style.FillGradientUnsupported ||
            !string.IsNullOrWhiteSpace(style.FillGradientStartColorArgb) ||
            !string.IsNullOrWhiteSpace(style.FillGradientEndColorArgb) ||
            style.FillGradientStops.Count > 0 ||
            style.FillGradientDegree.HasValue;

        private static ExcelCellBorderSnapshot? CreateTableBorder(string? colorArgb) {
            if (string.IsNullOrWhiteSpace(colorArgb)) {
                return null;
            }

            return new ExcelCellBorderSnapshot(
                left: new ExcelBorderSideSnapshot("thin", colorArgb),
                right: new ExcelBorderSideSnapshot("thin", colorArgb),
                top: new ExcelBorderSideSnapshot("thin", colorArgb),
                bottom: new ExcelBorderSideSnapshot("thin", colorArgb));
        }

    }
}
