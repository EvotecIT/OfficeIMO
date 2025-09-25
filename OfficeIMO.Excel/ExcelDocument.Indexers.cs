namespace OfficeIMO.Excel {
    /// <summary>
    /// Indexer helpers for convenient sheet access (name or 0-based index).
    /// </summary>
    public partial class ExcelDocument {
        /// <summary>
        /// Gets a worksheet by name (case-insensitive).
        /// </summary>
        public ExcelSheet this[string sheetName] {
            get {
                if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName));
                var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet is null) throw new ArgumentOutOfRangeException(nameof(sheetName), $"Sheet '{sheetName}' not found.");
                return sheet;
            }
        }

        /// <summary>
        /// Gets a worksheet by 0-based index in workbook order.
        /// </summary>
        public ExcelSheet this[int sheetIndex] {
            get {
                if (sheetIndex < 0 || sheetIndex >= Sheets.Count)
                    throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"Index {sheetIndex} is out of range (0..{Sheets.Count - 1}).");
                return Sheets[sheetIndex];
            }
        }
    }
}

