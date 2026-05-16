using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Lightweight object model wrapper for a single worksheet cell.
    /// </summary>
    public sealed class ExcelCell {
        internal ExcelCell(ExcelSheet sheet, int row, int column) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            Row = row;
            Column = column;
        }

        /// <summary>Gets the worksheet that owns the cell.</summary>
        public ExcelSheet Sheet { get; }
        /// <summary>Gets the 1-based row index.</summary>
        public int Row { get; }
        /// <summary>Gets the 1-based column index.</summary>
        public int Column { get; }
        /// <summary>Gets the A1 cell address.</summary>
        public string Address => A1.CellReference(Row, Column);

        /// <summary>
        /// Gets a typed snapshot of the cell value.
        /// </summary>
        public ExcelCellData GetValue() => Sheet.GetCellValueSnapshot(Row, Column);

        /// <summary>
        /// Gets formatted display text for the cell value.
        /// </summary>
        public string GetFormattedText(IFormatProvider? provider = null) => Sheet.GetCellFormattedText(Row, Column, provider);

        /// <summary>
        /// Gets the cell value converted to the requested type.
        /// </summary>
        public T? GetValue<T>() {
            object? value = GetValue().Value;
            if (value == null) {
                return default;
            }

            if (value is T typed) {
                return typed;
            }

            return (T?)Convert.ChangeType(value, typeof(T), System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Tries to get the cell value converted to the requested type.
        /// </summary>
        public bool TryGetValue<T>(out T? value) {
            try {
                value = GetValue<T>();
                return true;
            } catch (InvalidCastException) {
                value = default;
                return false;
            } catch (FormatException) {
                value = default;
                return false;
            } catch (OverflowException) {
                value = default;
                return false;
            } catch (ArgumentException) {
                value = default;
                return false;
            }
        }

        /// <summary>
        /// Sets the cell value.
        /// </summary>
        public ExcelCell SetValue(object? value) {
            Sheet.CellValue(Row, Column, value);
            return this;
        }

        /// <summary>
        /// Sets the cell formula.
        /// </summary>
        public ExcelCell SetFormula(string formula) {
            Sheet.CellFormula(Row, Column, formula);
            return this;
        }

        /// <summary>
        /// Clears selected cell data and metadata.
        /// </summary>
        public ExcelCell Clear(ExcelClearOptions options = ExcelClearOptions.All) {
            Sheet.ClearRange(Address + ":" + Address, options);
            return this;
        }

        /// <summary>
        /// Applies a number format to the cell.
        /// </summary>
        public ExcelCell SetNumberFormat(string numberFormat) {
            Sheet.FormatCell(Row, Column, numberFormat);
            return this;
        }

        /// <summary>
        /// Sets or clears bold font style.
        /// </summary>
        public ExcelCell SetBold(bool bold = true) {
            Sheet.CellBold(Row, Column, bold);
            return this;
        }

        /// <summary>
        /// Sets or clears italic font style.
        /// </summary>
        public ExcelCell SetItalic(bool italic = true) {
            Sheet.CellItalic(Row, Column, italic);
            return this;
        }

        /// <summary>
        /// Sets or clears underline font style.
        /// </summary>
        public ExcelCell SetUnderline(bool underline = true) {
            Sheet.CellUnderline(Row, Column, underline);
            return this;
        }

        /// <summary>
        /// Sets the font color using a hex color value.
        /// </summary>
        public ExcelCell SetFontColor(string hexColor) {
            Sheet.CellFontColor(Row, Column, hexColor);
            return this;
        }

        /// <summary>
        /// Sets the fill color using a hex color value.
        /// </summary>
        public ExcelCell SetFillColor(string hexColor) {
            Sheet.CellBackground(Row, Column, hexColor);
            return this;
        }

        /// <summary>
        /// Applies a border style to the cell.
        /// </summary>
        public ExcelCell SetBorder(BorderStyleValues style, string? hexColor = null) {
            Sheet.CellBorder(Row, Column, style, hexColor);
            return this;
        }

        /// <summary>
        /// Replaces the cell contents with rich inline text runs.
        /// </summary>
        public ExcelCell SetRichText(params ExcelRichTextRun[] runs) {
            Sheet.SetRichText(Row, Column, runs);
            return this;
        }

        /// <summary>
        /// Gets rich inline text runs from the cell.
        /// </summary>
        public IReadOnlyList<ExcelRichTextRun> GetRichText() => Sheet.GetRichText(Row, Column);
    }

    /// <summary>
    /// Lightweight object model wrapper for an A1 range.
    /// </summary>
    public sealed class ExcelRange {
        internal ExcelRange(ExcelSheet sheet, string address) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentNullException(nameof(address));

            var bounds = ParseRangeOrCell(address);
            Address = ToRangeAddress(bounds.r1, bounds.c1, bounds.r2, bounds.c2);
            FirstRow = bounds.r1;
            FirstColumn = bounds.c1;
            LastRow = bounds.r2;
            LastColumn = bounds.c2;
        }

        /// <summary>Gets the worksheet that owns the range.</summary>
        public ExcelSheet Sheet { get; }
        /// <summary>Gets the A1 range address.</summary>
        public string Address { get; }
        /// <summary>Gets the first row in the range.</summary>
        public int FirstRow { get; }
        /// <summary>Gets the first column in the range.</summary>
        public int FirstColumn { get; }
        /// <summary>Gets the last row in the range.</summary>
        public int LastRow { get; }
        /// <summary>Gets the last column in the range.</summary>
        public int LastColumn { get; }

        /// <summary>
        /// Gets a wrapper for the top-left cell.
        /// </summary>
        public ExcelCell FirstCell => Sheet.CellAt(FirstRow, FirstColumn);

        /// <summary>
        /// Clears selected data and metadata from the range.
        /// </summary>
        public ExcelRange Clear(ExcelClearOptions options = ExcelClearOptions.All) {
            Sheet.ClearRange(Address, options);
            return this;
        }

        /// <summary>
        /// Sorts the range by a 1-based column offset.
        /// </summary>
        public ExcelRange SortByColumn(int columnOffset, bool ascending = true, bool hasHeader = true) {
            Sheet.SortRangeByColumn(Address, columnOffset, ascending, hasHeader);
            return this;
        }

        /// <summary>
        /// Applies AutoFilter to the range using optional zero-based column criteria.
        /// </summary>
        public ExcelRange ApplyAutoFilter(Dictionary<uint, IEnumerable<string>>? filterCriteria = null) {
            Sheet.AddAutoFilter(Address, filterCriteria);
            return this;
        }

        /// <summary>
        /// Clears the worksheet AutoFilter.
        /// </summary>
        public ExcelRange ClearAutoFilter() {
            Sheet.AutoFilterClear();
            return this;
        }

        /// <summary>
        /// Merges the range.
        /// </summary>
        public ExcelRange Merge() {
            Sheet.MergeRange(Address);
            return this;
        }

        /// <summary>
        /// Removes merge definitions that overlap the range.
        /// </summary>
        public ExcelRange Unmerge() {
            Sheet.UnmergeRange(Address);
            return this;
        }

        /// <summary>
        /// Creates an Excel table over the range.
        /// </summary>
        public ExcelTable CreateTable(string name, bool hasHeader = true, TableStyle style = TableStyle.TableStyleMedium2, bool includeAutoFilter = true) {
            string resolvedName = Sheet.AddTableAndGetName(Address, hasHeader, name, style, includeAutoFilter);
            return Sheet.Table(resolvedName);
        }

        /// <summary>
        /// Applies a number format to every cell in the range.
        /// </summary>
        public ExcelRange SetNumberFormat(string numberFormat) {
            Sheet.FormatRange(Address, numberFormat);
            return this;
        }

        /// <summary>
        /// Applies a fill color to every cell in the range.
        /// </summary>
        public ExcelRange SetFillColor(string hexColor) {
            Sheet.FillRange(Address, hexColor);
            return this;
        }

        private static (int r1, int c1, int r2, int c2) ParseRangeOrCell(string address) {
            if (A1.TryParseRange(address, out int r1, out int c1, out int r2, out int c2)) {
                return (r1, c1, r2, c2);
            }

            var cell = A1.ParseCellRef(address);
            if (cell.Row <= 0 || cell.Col <= 0) {
                throw new ArgumentException($"Invalid A1 range or cell reference '{address}'.", nameof(address));
            }

            return (cell.Row, cell.Col, cell.Row, cell.Col);
        }

        private static string ToRangeAddress(int r1, int c1, int r2, int c2) {
            string start = A1.CellReference(r1, c1);
            string end = A1.CellReference(r2, c2);
            return $"{start}:{end}";
        }
    }

    /// <summary>
    /// Lightweight object model wrapper for an Excel table.
    /// </summary>
    public sealed class ExcelTable {
        internal ExcelTable(ExcelSheet sheet, string nameOrRange) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            NameOrRange = string.IsNullOrWhiteSpace(nameOrRange) ? throw new ArgumentNullException(nameof(nameOrRange)) : nameOrRange;
        }

        /// <summary>Gets the worksheet that owns the table.</summary>
        public ExcelSheet Sheet { get; }
        /// <summary>Gets the table name, display name, or A1 range used to locate the table.</summary>
        public string NameOrRange { get; }
        /// <summary>Gets the table range when it can be resolved.</summary>
        public string? Range => Sheet.GetTableRange(NameOrRange) ?? (A1.TryParseRange(NameOrRange, out _, out _, out _, out _) ? NameOrRange : null);

        /// <summary>
        /// Returns the table as a range wrapper.
        /// </summary>
        public ExcelRange AsRange() {
            string? range = Range;
            if (range == null) {
                throw new InvalidOperationException($"Table '{NameOrRange}' was not found on worksheet '{Sheet.Name}'.");
            }

            return Sheet.Range(range);
        }

        /// <summary>
        /// Applies a built-in table style and optional style flags.
        /// </summary>
        public ExcelTable SetStyle(TableStyle style, bool? showFirstColumn = null, bool? showLastColumn = null, bool? showRowStripes = null, bool? showColumnStripes = null) {
            Sheet.SetTableStyle(NameOrRange, style, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
            return this;
        }

        /// <summary>
        /// Applies totals row functions by header name.
        /// </summary>
        public ExcelTable SetTotals(IDictionary<string, TotalsRowFunctionValues> byHeader) {
            Sheet.SetTableTotalsByName(NameOrRange, byHeader);
            return this;
        }

        /// <summary>
        /// Clears totals row settings from the table.
        /// </summary>
        public ExcelTable ClearTotals() {
            Sheet.ClearTableTotals(NameOrRange);
            return this;
        }

        /// <summary>
        /// Appends rows from a data table to the Excel table.
        /// </summary>
        public ExcelTable AppendDataTable(System.Data.DataTable table) {
            Sheet.AppendDataTableToTable(table, NameOrRange);
            return this;
        }

        /// <summary>
        /// Sorts table rows by a 1-based column offset.
        /// </summary>
        public ExcelTable SortByColumn(int columnOffset, bool ascending = true) {
            AsRange().SortByColumn(columnOffset, ascending, hasHeader: true);
            return this;
        }
    }

    /// <summary>
    /// Describes a run of rich text inside a cell.
    /// </summary>
    public sealed class ExcelRichTextRun {
        /// <summary>
        /// Creates a rich text run with the supplied text.
        /// </summary>
        public ExcelRichTextRun(string text) {
            Text = text ?? string.Empty;
        }

        /// <summary>Gets or sets the run text.</summary>
        public string Text { get; set; }
        /// <summary>Gets or sets whether the run is bold.</summary>
        public bool Bold { get; set; }
        /// <summary>Gets or sets whether the run is italic.</summary>
        public bool Italic { get; set; }
        /// <summary>Gets or sets whether the run is underlined.</summary>
        public bool Underline { get; set; }
        /// <summary>Gets or sets the run font color as a hex value.</summary>
        public string? FontColor { get; set; }
        /// <summary>Gets or sets the run font name.</summary>
        public string? FontName { get; set; }
        /// <summary>Gets or sets the run font size.</summary>
        public double? FontSize { get; set; }

        /// <summary>
        /// Creates a plain rich text run.
        /// </summary>
        public static ExcelRichTextRun Plain(string text) => new ExcelRichTextRun(text);
    }
}
