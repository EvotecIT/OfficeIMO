using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Reader for an Excel workbook (read-only). Provides access to sheet readers and basic metadata.
    /// </summary>
    public sealed partial class ExcelDocumentReader : IDisposable
    {
        private readonly SpreadsheetDocument _doc;
        private readonly ExcelReadOptions _opt;
        private readonly SharedStringCache _sst;
        private readonly StylesCache _styles;

        private ExcelDocumentReader(SpreadsheetDocument doc, ExcelReadOptions opt)
        {
            _doc = doc;
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
            return new ExcelDocumentReader(doc, options ?? new ExcelReadOptions());
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
        public void Dispose() => _doc.Dispose();
    }
}
