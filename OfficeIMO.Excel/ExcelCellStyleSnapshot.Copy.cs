namespace OfficeIMO.Excel {
    public sealed partial class ExcelCellStyleSnapshot {
        /// <summary>Copies resolved visual metadata while replacing its immutable border snapshot.</summary>
        internal ExcelCellStyleSnapshot CopyWithBorder(ExcelCellBorderSnapshot? border) {
            var copy = (ExcelCellStyleSnapshot)MemberwiseClone();
            copy.Border = border;
            return copy;
        }
    }
}
