using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Reader for an Excel workbook (read-only). Provides access to sheet readers and basic metadata.
    /// </summary>
    public sealed partial class ExcelDocumentReader : IDisposable
    {
        private readonly SpreadsheetDocument _doc;
        private readonly bool _owns;
        private readonly ExcelReadOptions _opt;
        private readonly SharedStringCache _sst;
        private readonly StylesCache _styles;

        private ExcelDocumentReader(SpreadsheetDocument doc, ExcelReadOptions opt, bool owns)
        {
            _doc = doc;
            _owns = owns;
            _opt = opt ?? new ExcelReadOptions();
            _sst = SharedStringCache.Build(doc);
            _styles = StylesCache.Build(doc);
        }

        /// <summary>
        /// Opens an Excel file for read-only access.
        /// </summary>
        public static ExcelDocumentReader Open(string path, ExcelReadOptions? options = null)
        {
            var doc = SpreadsheetDocument.Open(path, false);
            return new ExcelDocumentReader(doc, options ?? new ExcelReadOptions(), owns: true);
        }

        /// <summary>
        /// Wraps an already open SpreadsheetDocument without taking ownership.
        /// The returned reader must be disposed, but it will not close the underlying document.
        /// </summary>
        public static ExcelDocumentReader Wrap(SpreadsheetDocument document, ExcelReadOptions? options = null)
        {
            if (document is null) throw new ArgumentNullException(nameof(document));
            return new ExcelDocumentReader(document, options ?? new ExcelReadOptions(), owns: false);
        }

        /// <summary>
        /// Returns the list of sheet names in workbook order.
        /// </summary>
        public IReadOnlyList<string> GetSheetNames()
        {
            var wb = _doc.WorkbookPart!.Workbook;
            return wb.Sheets!.Elements<Sheet>().Select(s => s.Name!.Value!).ToList();
        }

        /// <summary>
        /// Gets a reader for the specified worksheet name.
        /// </summary>
        public ExcelSheetReader GetSheet(string name)
        {
            var wb = _doc.WorkbookPart!.Workbook;
            var sheet = wb.Sheets!.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal));
            if (sheet is null) throw new KeyNotFoundException($"Sheet '{name}' not found.");
            var wsPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheet.Id!);
            return new ExcelSheetReader(name, wsPart, _sst, _styles, _opt);
        }

        /// <summary>
        /// Disposes the underlying OpenXML document.
        /// </summary>
        public void Dispose()
        {
            if (_owns)
                _doc.Dispose();
        }
    }
}
