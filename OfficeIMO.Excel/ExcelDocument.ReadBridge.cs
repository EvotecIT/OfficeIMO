using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel.Read;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Convenience bridge methods exposing read helpers on an open ExcelDocument to avoid re-opening files.
    /// </summary>
    public partial class ExcelDocument
    {
        /// <summary>
        /// Creates a reader that shares this document's underlying OpenXML handle (no new file handle).
        /// Caller should dispose the reader after use; it will not close this document.
        /// </summary>
        public ExcelDocumentReader CreateReader(ExcelReadOptions? options = null)
            => ExcelDocumentReader.Wrap(_spreadSheetDocument, options ?? new ExcelReadOptions());

        /// <summary>
        /// Returns worksheet names in workbook order.
        /// </summary>
        public IReadOnlyList<string> GetSheetNames()
            => Sheets.Select(s => s.Name).ToList();

        /// <summary>
        /// Gets a worksheet by name.
        /// </summary>
        public ExcelSheet GetSheet(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));
            var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal));
            if (sheet is null) throw new KeyNotFoundException($"Sheet '{name}' not found.");
            return sheet;
        }

        /// <summary>
        /// Tries to get a worksheet by name.
        /// </summary>
        public bool TryGetSheet(string name, out ExcelSheet sheet)
        {
            sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal));
            return sheet != null;
        }
    }
}

