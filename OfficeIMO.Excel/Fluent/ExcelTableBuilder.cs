namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent helpers for adding Excel tables over a range.
    /// </summary>
    public class ExcelTableBuilder {
        private readonly ExcelSheet _sheet;
        private ExcelTableStyle _style = ExcelTableStyle.TableStyleLight9;
        private bool _hasHeader = true;
        private bool _includeAutoFilter = true;

        internal ExcelTableBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        /// <summary>Sets the table style to use when adding tables.</summary>
        public ExcelTableBuilder Style(ExcelTableStyle style) {
            _style = style;
            return this;
        }

        /// <summary>Specifies whether the first row of the range is a header row.</summary>
        public ExcelTableBuilder HasHeader(bool hasHeader = true) {
            _hasHeader = hasHeader;
            return this;
        }

        /// <summary>Specifies whether to enable AutoFilter for the table.</summary>
        public ExcelTableBuilder WithAutoFilter(bool includeAutoFilter = true) {
            _includeAutoFilter = includeAutoFilter;
            return this;
        }

        internal void Build(string range, string name) {
            _sheet.AddTable(range, _hasHeader, name, _style, _includeAutoFilter);
        }

        /// <summary>
        /// Adds a table over the specified A1 <paramref name="range"/> with optional header, name, style and AutoFilter.
        /// </summary>
        public ExcelTableBuilder Add(string range, bool hasHeader = true, string name = "", ExcelTableStyle style = ExcelTableStyle.TableStyleLight9, bool includeAutoFilter = true) {
            _sheet.AddTable(range, hasHeader, name, style, includeAutoFilter);
            return this;
        }
    }
}
