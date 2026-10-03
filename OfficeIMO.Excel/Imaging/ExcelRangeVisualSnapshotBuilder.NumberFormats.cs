namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeVisualSnapshotBuilder {
        private static void ApplyNumericFormatColor(ExcelCellStyleSnapshot style, ExcelCellData data) {
            if (data.Kind is not ExcelCellDataKind.Number and not ExcelCellDataKind.Formula
                || data.Value is not double value) return;
            string? color = ExcelNumberFormatDisplay.GetNumericFormatColor(value,
                style.NumberFormatId, style.NumberFormatCode);
            // Each cell owns a fresh style snapshot. Conditional formatting is
            // applied later and can override this selected number-format color.
            if (color != null) style.FontColorArgb = "FF" + color;
        }
    }
}
