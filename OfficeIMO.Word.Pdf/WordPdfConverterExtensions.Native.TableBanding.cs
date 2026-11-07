namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Use the source table's body range and band sizes for every projected property.
        // PDF presentation stripes have a fixed phase and cannot represent Word band groups.
        private static NativeTableConditionalStyleDefaults GetNativeHorizontalBandStyle(WordTable table,
            NativeTableStyleDefaults defaults, int row, int headerRows, int footerStart) {
            if (table.ConditionalFormattingNoHorizontalBand == true || defaults.RowBandSize <= 0 ||
                row < headerRows || row >= footerStart) return NativeTableConditionalStyleDefaults.Empty;
            return (row - headerRows) / defaults.RowBandSize % 2 == 0
                ? defaults.Band1HorizontalStyle : defaults.Band2HorizontalStyle;
        }

        private static NativeTableConditionalStyleDefaults GetNativeVerticalBandStyle(WordTable table,
            NativeTableStyleDefaults defaults, int row, int column, int headerRows, int footerStart) {
            int bandColumn = column - (table.ConditionalFormattingFirstColumn == true ? 1 : 0);
            if (table.ConditionalFormattingNoVerticalBand == true || defaults.ColumnBandSize <= 0 ||
                bandColumn < 0 || row < headerRows || row >= footerStart) return NativeTableConditionalStyleDefaults.Empty;
            return bandColumn / defaults.ColumnBandSize % 2 == 0
                ? defaults.Band1VerticalStyle : defaults.Band2VerticalStyle;
        }

    }
}
