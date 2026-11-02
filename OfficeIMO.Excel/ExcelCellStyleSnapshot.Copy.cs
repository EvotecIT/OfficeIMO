namespace OfficeIMO.Excel {
    public sealed partial class ExcelCellStyleInfo {
        /// <summary>Copies resolved visual metadata while replacing its immutable border snapshot.</summary>
        internal ExcelCellStyleInfo CopyWithBorder(ExcelCellBorderInfo? border) {
            var copy = (ExcelCellStyleInfo)MemberwiseClone();
            copy.Border = border;
            return copy;
        }
    }
}
