using System;
using System.Collections.Generic;
using OfficeIMO.Excel.Read;
using OfficeIMO.Excel.Read.Edit;

namespace OfficeIMO.Excel.Read.Fluent
{
    /// <summary>
    /// Entry point for fluent read pipelines.
    /// </summary>
    public sealed class ExcelFluentReadWorkbook
    {
        private readonly ExcelDocument _doc;
        private readonly ExcelReadOptions _options;

        internal ExcelFluentReadWorkbook(ExcelDocument doc, ExcelReadOptions? options)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _options = options ?? new ExcelReadOptions();
        }

        public ExcelFluentReadSheet Sheet(string name)
            => new ExcelFluentReadSheet(_doc, name, _options);

        public ExcelFluentReadSheet Sheet(int index)
        {
            if (index < 0 || index >= _doc.Sheets.Count) throw new ArgumentOutOfRangeException(nameof(index));
            return new ExcelFluentReadSheet(_doc, _doc.Sheets[index].Name, _options);
        }
    }

    /// <summary>
    /// Fluent sheet scope.
    /// </summary>
    public sealed class ExcelFluentReadSheet
    {
        private readonly ExcelDocument _doc;
        private readonly string _sheetName;
        private readonly ExcelReadOptions _options;

        internal ExcelFluentReadSheet(ExcelDocument doc, string sheetName, ExcelReadOptions options)
        {
            _doc = doc;
            _sheetName = sheetName;
            _options = options;
        }

        public ExcelFluentReadRange UsedRange()
        {
            var sheet = _doc[_sheetName];
            var a1 = sheet.UsedRangeA1;
            return new ExcelFluentReadRange(_doc, _sheetName, a1, _options);
        }

        public ExcelFluentReadRange Range(string a1Range)
            => new ExcelFluentReadRange(_doc, _sheetName, a1Range, _options);
    }

    /// <summary>
    /// Fluent range scope and materializers.
    /// </summary>
    public sealed class ExcelFluentReadRange
    {
        private readonly ExcelDocument _doc;
        private readonly string _sheetName;
        private readonly string _a1;
        private readonly ExcelReadOptions _options;

        internal ExcelFluentReadRange(ExcelDocument doc, string sheetName, string a1Range, ExcelReadOptions options)
        {
            _doc = doc;
            _sheetName = sheetName;
            _a1 = a1Range;
            _options = options;
        }

        /// <summary>
        /// Prefer decimals when converting numeric cells.
        /// </summary>
        public ExcelFluentReadRange NumericAsDecimal(bool enable = true)
        {
            _options.NumericAsDecimal = enable;
            return this;
        }

        /// <summary>
        /// Reads the range as a sequence of dictionaries using the first row as headers.
        /// </summary>
        public IEnumerable<System.Collections.Generic.Dictionary<string, object?>> AsRows()
        {
            using var rdr = ExcelDocumentReader.Wrap(_doc._spreadSheetDocument, _options);
            return rdr.GetSheet(_sheetName).ReadObjects(_a1);
        }

        /// <summary>
        /// Maps rows (excluding header) to instances of T by matching headers to property names.
        /// </summary>
        public IEnumerable<T> AsObjects<T>() where T : new()
        {
            using var rdr = ExcelDocumentReader.Wrap(_doc._spreadSheetDocument, _options);
            return rdr.GetSheet(_sheetName).ReadObjects<T>(_a1);
        }

        /// <summary>
        /// Returns editable rows over the range (header-aware cell access + write-back).
        /// </summary>
        public IEnumerable<RowEdit> AsEditableRows()
        {
            var sheet = _doc[_sheetName];
            return sheet.RowsObjects(_a1, _options);
        }
    }
}
