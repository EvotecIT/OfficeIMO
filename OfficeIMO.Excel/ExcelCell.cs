using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Lightweight object model wrapper for a single worksheet cell.
    /// </summary>
    public sealed partial class ExcelCell {
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
        /// Sets the native Excel underline style.
        /// </summary>
        public ExcelCell SetUnderline(ExcelUnderlineStyle underlineStyle) {
            Sheet.CellUnderline(Row, Column, underlineStyle);
            return this;
        }

        /// <summary>
        /// Sets or clears strikethrough font style.
        /// </summary>
        public ExcelCell SetStrikethrough(bool strikethrough = true) {
            Sheet.CellStrikethrough(Row, Column, strikethrough);
            return this;
        }

        /// <summary>
        /// Sets the native Excel baseline, superscript, or subscript alignment.
        /// </summary>
        public ExcelCell SetVerticalTextAlignment(ExcelVerticalTextAlignment alignment) {
            Sheet.CellVerticalTextAlignment(Row, Column, alignment);
            return this;
        }

        /// <summary>Formats the cell text as superscript.</summary>
        public ExcelCell SetSuperscript() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Superscript);

        /// <summary>Formats the cell text as subscript.</summary>
        public ExcelCell SetSubscript() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Subscript);

        /// <summary>Restores the cell text to the normal baseline.</summary>
        public ExcelCell SetBaseline() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Baseline);

        /// <summary>
        /// Changes stored text casing while preserving cell or rich-run formatting.
        /// Formulas and non-text values are left unchanged.
        /// </summary>
        public ExcelCell TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            Sheet.TransformCellTextCase(Row, Column, textCase, culture);
            return this;
        }

        /// <summary>
        /// Sets the font family name.
        /// </summary>
        public ExcelCell SetFontName(string fontName) {
            Sheet.CellFontName(Row, Column, fontName);
            return this;
        }

        /// <summary>
        /// Sets the font size in points.
        /// </summary>
        public ExcelCell SetFontSize(double fontSize) {
            Sheet.CellFontSize(Row, Column, fontSize);
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
        /// Sets or clears shrink-to-fit text alignment.
        /// </summary>
        public ExcelCell SetShrinkToFit(bool shrinkToFit = true) {
            Sheet.CellShrinkToFit(Row, Column, shrinkToFit);
            return this;
        }

        /// <summary>
        /// Sets Excel text rotation. Use 0-90 for upward rotation, 91-180 for downward rotation, or 255 for stacked vertical text.
        /// </summary>
        public ExcelCell SetTextRotation(int rotation) {
            Sheet.CellTextRotation(Row, Column, rotation);
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
        /// Sets a two-color linear gradient fill using hex color values.
        /// </summary>
        public ExcelCell SetGradientFill(string fromHexColor, string toHexColor, double degree = 0) {
            Sheet.CellGradientBackground(Row, Column, fromHexColor, toHexColor, degree);
            return this;
        }

        /// <summary>
        /// Applies a border style to the cell.
        /// </summary>
        public ExcelCell SetBorder(ExcelBorderStyle style, string? hexColor = null) {
            Sheet.CellBorder(Row, Column, style, hexColor);
            return this;
        }

        /// <summary>
        /// Applies diagonal borders to the cell.
        /// </summary>
        public ExcelCell SetDiagonalBorder(ExcelBorderStyle style, string? hexColor = null, bool diagonalUp = true, bool diagonalDown = true) {
            Sheet.CellDiagonalBorder(Row, Column, style, hexColor, diagonalUp, diagonalDown);
            return this;
        }

        /// <summary>Applies a decimal number format.</summary>
        public ExcelCell Number(int decimals = 2) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals));

        /// <summary>Applies a whole-number format with thousands separators.</summary>
        public ExcelCell Integer() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Integer));

        /// <summary>Applies a percent number format.</summary>
        public ExcelCell Percent(int decimals = 0) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals));

        /// <summary>Applies a currency number format.</summary>
        public ExcelCell Currency(int decimals = 2, CultureInfo? culture = null) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals, culture));

        /// <summary>Applies a date number format.</summary>
        public ExcelCell Date(string pattern = "yyyy-mm-dd") => SetNumberFormat(pattern);

        /// <summary>Applies a date/time number format.</summary>
        public ExcelCell DateTime(string pattern = "yyyy-mm-dd hh:mm:ss") => SetNumberFormat(pattern);

        /// <summary>Applies a time number format.</summary>
        public ExcelCell Time() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Time));

        /// <summary>Applies an elapsed-hours duration format.</summary>
        public ExcelCell DurationHours() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours));

        /// <summary>Applies a text number format.</summary>
        public ExcelCell Text() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Text));

        /// <summary>Applies a positive/success status style.</summary>
        public ExcelCell Success() => SetFillColor("E7F6E7").SetFontColor("226B22");

        /// <summary>Applies a warning status style.</summary>
        public ExcelCell Warning() => SetFillColor("FFF4CC").SetFontColor("7A4D00");

        /// <summary>Applies an error status style.</summary>
        public ExcelCell Error() => SetFillColor("FCE4E4").SetFontColor("9C0006");

        /// <summary>Applies a muted text style.</summary>
        public ExcelCell MutedText() => SetFontColor("666666");

        /// <summary>Applies a simple report header style.</summary>
        public ExcelCell HeaderStyle() => SetBold().SetFillColor("D9EAF7").SetFontColor("1F4E79");

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
    public sealed partial class ExcelRange {
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
        /// Builds data validation rules for the range.
        /// </summary>
        public ExcelRangeDataValidationBuilder Validation => new ExcelRangeDataValidationBuilder(this);

        /// <summary>
        /// Builds data validation rules for the range.
        /// </summary>
        public ExcelRangeDataValidationBuilder Validate => Validation;

        /// <summary>
        /// Builds conditional formatting rules for the range.
        /// </summary>
        public ExcelRangeConditionalFormattingBuilder ConditionalFormatting => new ExcelRangeConditionalFormattingBuilder(this);

        /// <summary>
        /// Builds conditional formatting rules for the range.
        /// </summary>
        public ExcelRangeConditionalFormattingBuilder ConditionalFormat => ConditionalFormatting;

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

        /// <summary>Transactionally inserts a cell block at this range.</summary>
        public ExcelMutationResult Insert(ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.InsertCells(Address, direction, options, cancellationToken);

        /// <summary>Transactionally deletes this cell block.</summary>
        public ExcelMutationResult Delete(ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.DeleteCells(Address, direction, options, cancellationToken);

        /// <summary>Copies this range to a destination top-left cell.</summary>
        public ExcelMutationResult CopyTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.CopyRange(Address, destinationTopLeft, options, cancellationToken);

        /// <summary>Moves this range to a destination top-left cell.</summary>
        public ExcelMutationResult MoveTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.MoveRange(Address, destinationTopLeft, options, cancellationToken);

        /// <summary>Copies this range to a destination with rows and columns transposed.</summary>
        public ExcelMutationResult TransposeTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.TransposeRange(Address, destinationTopLeft, options, cancellationToken);

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
        public ExcelTable CreateTable(string name, bool hasHeader = true, ExcelTableStyle style = ExcelTableStyle.TableStyleMedium2, bool includeAutoFilter = true) {
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

        /// <summary>
        /// Applies a two-color linear gradient fill to every cell in the range.
        /// </summary>
        public ExcelRange SetGradientFill(string fromHexColor, string toHexColor, double degree = 0) {
            Sheet.FillRangeGradient(Address, fromHexColor, toHexColor, degree);
            return this;
        }

        /// <summary>
        /// Applies a font color to every cell in the range.
        /// </summary>
        public ExcelRange SetFontColor(string hexColor) {
            ForEachCell((row, column) => Sheet.CellFontColor(row, column, hexColor));
            return this;
        }

        /// <summary>
        /// Applies a font family name to every cell in the range.
        /// </summary>
        public ExcelRange SetFontName(string fontName) {
            ForEachCell((row, column) => Sheet.CellFontName(row, column, fontName));
            return this;
        }

        /// <summary>
        /// Applies a font size in points to every cell in the range.
        /// </summary>
        public ExcelRange SetFontSize(double fontSize) {
            ForEachCell((row, column) => Sheet.CellFontSize(row, column, fontSize));
            return this;
        }

        /// <summary>
        /// Sets or clears bold font style for every cell in the range.
        /// </summary>
        public ExcelRange SetBold(bool bold = true) {
            ForEachCell((row, column) => Sheet.CellBold(row, column, bold));
            return this;
        }

        /// <summary>Sets or clears italic font style for every cell in the range.</summary>
        public ExcelRange SetItalic(bool italic = true) {
            ForEachCell((row, column) => Sheet.CellItalic(row, column, italic));
            return this;
        }

        /// <summary>Sets or clears single underline font style for every cell in the range.</summary>
        public ExcelRange SetUnderline(bool underline = true) {
            ForEachCell((row, column) => Sheet.CellUnderline(row, column, underline));
            return this;
        }

        /// <summary>Sets the native Excel underline style for every cell in the range.</summary>
        public ExcelRange SetUnderline(ExcelUnderlineStyle underlineStyle) {
            ForEachCell((row, column) => Sheet.CellUnderline(row, column, underlineStyle));
            return this;
        }

        /// <summary>Sets or clears strikethrough font style for every cell in the range.</summary>
        public ExcelRange SetStrikethrough(bool strikethrough = true) {
            ForEachCell((row, column) => Sheet.CellStrikethrough(row, column, strikethrough));
            return this;
        }

        /// <summary>Sets the native Excel baseline, superscript, or subscript alignment for every cell in the range.</summary>
        public ExcelRange SetVerticalTextAlignment(ExcelVerticalTextAlignment alignment) {
            ForEachCell((row, column) => Sheet.CellVerticalTextAlignment(row, column, alignment));
            return this;
        }

        /// <summary>Formats every cell in the range as superscript.</summary>
        public ExcelRange SetSuperscript() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Superscript);

        /// <summary>Formats every cell in the range as subscript.</summary>
        public ExcelRange SetSubscript() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Subscript);

        /// <summary>Restores every cell in the range to the normal baseline.</summary>
        public ExcelRange SetBaseline() => SetVerticalTextAlignment(ExcelVerticalTextAlignment.Baseline);

        /// <summary>
        /// Changes stored text casing in text cells while preserving cell and rich-run formatting.
        /// Formulas and non-text values are left unchanged.
        /// </summary>
        public ExcelRange TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            ForEachCell((row, column) => Sheet.TransformCellTextCase(row, column, textCase, culture));
            return this;
        }

        /// <summary>
        /// Sets or clears shrink-to-fit text alignment for every cell in the range.
        /// </summary>
        public ExcelRange SetShrinkToFit(bool shrinkToFit = true) {
            ForEachCell((row, column) => Sheet.CellShrinkToFit(row, column, shrinkToFit));
            return this;
        }

        /// <summary>
        /// Sets Excel text rotation for every cell in the range.
        /// </summary>
        public ExcelRange SetTextRotation(int rotation) {
            ForEachCell((row, column) => Sheet.CellTextRotation(row, column, rotation));
            return this;
        }

        /// <summary>Applies a decimal number format.</summary>
        public ExcelRange Number(int decimals = 2) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals));

        /// <summary>Applies a whole-number format with thousands separators.</summary>
        public ExcelRange Integer() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Integer));

        /// <summary>Applies a percent number format.</summary>
        public ExcelRange Percent(int decimals = 0) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals));

        /// <summary>Applies a currency number format.</summary>
        public ExcelRange Currency(int decimals = 2, CultureInfo? culture = null) => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals, culture));

        /// <summary>Applies a date number format.</summary>
        public ExcelRange Date(string pattern = "yyyy-mm-dd") => SetNumberFormat(pattern);

        /// <summary>Applies a date/time number format.</summary>
        public ExcelRange DateTime(string pattern = "yyyy-mm-dd hh:mm:ss") => SetNumberFormat(pattern);

        /// <summary>Applies a time number format.</summary>
        public ExcelRange Time() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Time));

        /// <summary>Applies an elapsed-hours duration format.</summary>
        public ExcelRange DurationHours() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours));

        /// <summary>Applies a text number format.</summary>
        public ExcelRange Text() => SetNumberFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Text));

        /// <summary>Applies a positive/success status style.</summary>
        public ExcelRange Success() => SetFillColor("E7F6E7").SetFontColor("226B22");

        /// <summary>Applies a warning status style.</summary>
        public ExcelRange Warning() => SetFillColor("FFF4CC").SetFontColor("7A4D00");

        /// <summary>Applies an error status style.</summary>
        public ExcelRange Error() => SetFillColor("FCE4E4").SetFontColor("9C0006");

        /// <summary>Applies a muted text style.</summary>
        public ExcelRange MutedText() => SetFontColor("666666");

        /// <summary>Applies a simple report header style.</summary>
        public ExcelRange HeaderStyle() => SetBold().SetFillColor("D9EAF7").SetFontColor("1F4E79");

        private void ForEachCell(Action<int, int> apply) {
            for (int row = FirstRow; row <= LastRow; row++) {
                for (int column = FirstColumn; column <= LastColumn; column++) {
                    apply(row, column);
                }
            }
        }

        private static (int r1, int c1, int r2, int c2) ParseRangeOrCell(string address) {
            string normalizedAddress = address.Replace("$", string.Empty);
            if (A1.TryParseRange(normalizedAddress, out int r1, out int c1, out int r2, out int c2)) {
                return (r1, c1, r2, c2);
            }

            var cell = A1.ParseCellRef(normalizedAddress);
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
        public ExcelTable SetStyle(ExcelTableStyle style, bool? showFirstColumn = null, bool? showLastColumn = null, bool? showRowStripes = null, bool? showColumnStripes = null) {
            Sheet.SetTableStyle(NameOrRange, style, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
            return this;
        }

        /// <summary>
        /// Applies totals row functions by header name.
        /// </summary>
        public ExcelTable SetTotals(IDictionary<string, ExcelTableTotalsFunction> byHeader) {
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

        /// <summary>Renames this table and its workbook references.</summary>
        public ExcelTable Rename(string newName, ExcelTableNameValidationMode validationMode = ExcelTableNameValidationMode.Strict) {
            return new ExcelTable(Sheet, Sheet.RenameTable(NameOrRange, newName, validationMode));
        }

        /// <summary>Replaces the table's ordered column schema and optionally resizes its range.</summary>
        public ExcelTable SetSchema(IReadOnlyList<string> columnNames, string? newRange = null) {
            string stableName = Sheet.ResolveTableName(NameOrRange);
            Sheet.SetTableSchema(stableName, columnNames, newRange);
            return new ExcelTable(Sheet, stableName);
        }

        /// <summary>Resizes the table, preserving current column names where possible.</summary>
        public ExcelTable Resize(string newRange) {
            string stableName = Sheet.ResolveTableName(NameOrRange);
            Sheet.ResizeTable(stableName, newRange);
            return new ExcelTable(Sheet, stableName);
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

        /// <summary>
        /// Resolves a data-column range in this table by its header text.
        /// </summary>
        public ExcelRange Column(string headerName, bool includeHeader = false, bool normalizeHeader = true) {
            string range = Sheet.GetColumnRangeByHeader(headerName, NameOrRange, headerRow: 0, includeHeader, normalizeHeader);
            return Sheet.Range(range);
        }
    }

    /// <summary>
    /// Describes a run of rich text inside a cell.
    /// </summary>
    public sealed class ExcelRichTextRun {
        private bool _bold;
        private bool _italic;
        private bool _underline;
        private bool _strikethrough;
        private ExcelUnderlineStyle? _underlineStyle;

        /// <summary>
        /// Creates a rich text run with the supplied text.
        /// </summary>
        public ExcelRichTextRun(string text) {
            Text = text ?? string.Empty;
        }

        /// <summary>Gets or sets the run text.</summary>
        public string Text { get; set; }
        /// <summary>Gets or sets whether the run is bold.</summary>
        public bool Bold {
            get => _bold;
            set {
                _bold = value;
                BoldSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is italic.</summary>
        public bool Italic {
            get => _italic;
            set {
                _italic = value;
                ItalicSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is underlined.</summary>
        public bool Underline {
            get => _underline;
            set {
                _underline = value;
                UnderlineSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is struck through.</summary>
        public bool Strikethrough {
            get => _strikethrough;
            set {
                _strikethrough = value;
                StrikethroughSpecified = true;
            }
        }
        /// <summary>Gets or sets the run underline style.</summary>
        public ExcelUnderlineStyle? UnderlineStyle {
            get => _underlineStyle;
            set {
                _underlineStyle = value;
                if (value.HasValue) UnderlineSpecified = true;
            }
        }
        /// <summary>Gets or sets the run font color as a hex value.</summary>
        public string? FontColor { get; set; }
        /// <summary>Gets or sets the run font name.</summary>
        public string? FontName { get; set; }
        /// <summary>Gets or sets the run font size.</summary>
        public double? FontSize { get; set; }
        /// <summary>Gets or sets the run vertical text alignment.</summary>
        public ExcelVerticalTextAlignment? VerticalTextAlignment { get; set; }
        /// <summary>Gets or sets whether the run font uses outline text.</summary>
        public bool Outline { get; set; }
        /// <summary>Gets or sets whether the run font uses shadow text.</summary>
        public bool Shadow { get; set; }
        /// <summary>Gets or sets whether the run font uses condensed text.</summary>
        public bool Condense { get; set; }
        /// <summary>Gets or sets whether the run font uses extended text.</summary>
        public bool Extend { get; set; }
        /// <summary>Gets or sets the run font family classification byte.</summary>
        public byte? FontFamily { get; set; }
        /// <summary>Gets or sets the run font character set byte.</summary>
        public byte? FontCharacterSet { get; set; }

        internal bool BoldSpecified { get; private set; }
        internal bool ItalicSpecified { get; private set; }
        internal bool UnderlineSpecified { get; private set; }
        internal bool StrikethroughSpecified { get; private set; }

        /// <summary>
        /// Creates a plain rich text run.
        /// </summary>
        public static ExcelRichTextRun Plain(string text) => new ExcelRichTextRun(text);

        /// <summary>
        /// Changes the stored run text casing while preserving rich-text formatting.
        /// </summary>
        public ExcelRichTextRun TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            Text = OfficeTextCaseTransformer.Apply(Text, textCase, culture);
            return this;
        }

        /// <summary>Projects native run properties without losing explicit disabled values.</summary>
        internal static ExcelRichTextRun FromOpenXml(string text, RunProperties? properties, string? resolvedFontColor = null) {
            var result = new ExcelRichTextRun(text) {
                FontColor = resolvedFontColor ?? properties?.GetFirstChild<Color>()?.Rgb?.Value,
                FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value,
                VerticalTextAlignment = GetVerticalTextAlignment(properties),
                Outline = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Outline>()),
                Shadow = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Shadow>()),
                Condense = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Condense>()),
                Extend = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Extend>()),
                FontFamily = GetFontFamily(properties),
                FontCharacterSet = GetFontCharacterSet(properties)
            };

            DocumentFormat.OpenXml.Spreadsheet.Bold? bold = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Bold>();
            if (bold != null) result.Bold = ExcelOpenXmlFontProperty.IsEnabled(bold);
            DocumentFormat.OpenXml.Spreadsheet.Italic? italic = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Italic>();
            if (italic != null) result.Italic = ExcelOpenXmlFontProperty.IsEnabled(italic);
            DocumentFormat.OpenXml.Spreadsheet.Underline? underline = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>();
            if (underline != null) {
                result.Underline = ExcelOpenXmlFontProperty.IsUnderlineEnabled(underline);
                result.UnderlineStyle = GetUnderlineStyle(properties);
            }
            Strike? strike = properties?.GetFirstChild<Strike>();
            if (strike != null) result.Strikethrough = ExcelOpenXmlFontProperty.IsEnabled(strike);
            return result;
        }

        /// <summary>Copies a run while retaining which direct Boolean properties were present.</summary>
        internal ExcelRichTextRun Clone(string? text = null) {
            var result = new ExcelRichTextRun(text ?? Text) {
                FontColor = FontColor,
                FontName = FontName,
                FontSize = FontSize,
                VerticalTextAlignment = VerticalTextAlignment,
                Outline = Outline,
                Shadow = Shadow,
                Condense = Condense,
                Extend = Extend,
                FontFamily = FontFamily,
                FontCharacterSet = FontCharacterSet
            };
            if (BoldSpecified) result.Bold = Bold;
            if (ItalicSpecified) result.Italic = Italic;
            if (UnderlineSpecified) {
                result.Underline = Underline;
                result.UnderlineStyle = UnderlineStyle;
            }
            if (StrikethroughSpecified) result.Strikethrough = Strikethrough;
            return result;
        }

        internal static void AppendFontMetadata(RunProperties properties, ExcelRichTextRun run) {
            if (run.VerticalTextAlignment.HasValue) {
                properties.Append(new VerticalTextAlignment { Val = run.VerticalTextAlignment.Value.ToOpenXml() });
            }

            if (run.Outline) {
                properties.Append(new Outline());
            }

            if (run.Shadow) {
                properties.Append(new Shadow());
            }

            if (run.Condense) {
                properties.Append(new Condense());
            }

            if (run.Extend) {
                properties.Append(new Extend());
            }

            if (run.FontFamily.HasValue) {
                properties.Append(new FontFamilyNumbering { Val = run.FontFamily.Value });
            }

            if (run.FontCharacterSet.HasValue) {
                var charset = new OpenXmlUnknownElement(string.Empty, "charset", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                charset.SetAttribute(new OpenXmlAttribute("val", string.Empty, run.FontCharacterSet.Value.ToString(CultureInfo.InvariantCulture)));
                properties.Append(charset);
            }
        }

        internal static byte? GetFontFamily(RunProperties? properties) {
            return TryGetByte(properties?.GetFirstChild<FontFamilyNumbering>()?.Val?.Value, out byte family)
                ? family
                : null;
        }

        internal static byte? GetFontCharacterSet(RunProperties? properties) {
            OpenXmlElement? charset = properties?.ChildElements.FirstOrDefault(child =>
                string.Equals(child.LocalName, "charset", StringComparison.Ordinal)
                && string.Equals(child.NamespaceUri, "http://schemas.openxmlformats.org/spreadsheetml/2006/main", StringComparison.Ordinal));
            return TryGetByte(charset?.GetAttribute("val", string.Empty).Value, out byte characterSet)
                ? characterSet
                : null;
        }

        internal static ExcelVerticalTextAlignment? GetVerticalTextAlignment(RunProperties? properties) {
            return properties?.GetFirstChild<VerticalTextAlignment>()?.Val?.Value.ToOfficeEnum();
        }

        internal static ExcelUnderlineStyle? GetUnderlineStyle(RunProperties? properties) {
            Underline? underline = properties?.GetFirstChild<Underline>();
            if (underline == null) {
                return null;
            }

            return underline.Val?.Value.ToOfficeEnum() ?? ExcelUnderlineStyle.Single;
        }

        private static bool TryGetByte(uint? value, out byte result) {
            if (value.HasValue && value.Value <= byte.MaxValue) {
                result = checked((byte)value.Value);
                return true;
            }

            result = 0;
            return false;
        }

        private static bool TryGetByte(int? value, out byte result) {
            if (value.HasValue && value.Value >= 0 && value.Value <= byte.MaxValue) {
                result = checked((byte)value.Value);
                return true;
            }

            result = 0;
            return false;
        }

        private static bool TryGetByte(string? value, out byte result) {
            if (!string.IsNullOrWhiteSpace(value)
                && uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint parsed)
                && parsed <= byte.MaxValue) {
                result = checked((byte)parsed);
                return true;
            }

            result = 0;
            return false;
        }
    }
}
