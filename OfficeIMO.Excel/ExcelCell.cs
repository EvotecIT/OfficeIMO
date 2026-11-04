using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OfficeIMO.Spreadsheet;
using OfficeIMO.Drawing;

// Cell and range setters mirror each other, so paired overloads with optional arguments are intended.
#pragma warning disable RS0026, RS0027
namespace OfficeIMO.Excel {
    /// <summary>
    /// A worksheet cell handle: the sheet plus a 1-based row and column. It stores no cell state, so creating one
    /// in a loop allocates nothing; every member reads or writes the worksheet directly.
    /// </summary>
    public readonly partial struct ExcelCell : IEquatable<ExcelCell> {
        private readonly ExcelSheet? _sheet;

        internal ExcelCell(ExcelSheet sheet, int row, int column) {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            Row = row;
            Column = column;
        }

        /// <summary>Gets the worksheet that owns the cell.</summary>
        /// <exception cref="InvalidOperationException">The handle is <c>default</c> and has no worksheet.</exception>
        public ExcelSheet Sheet => _sheet ?? throw new InvalidOperationException("The cell handle is not attached to a worksheet. Use ExcelSheet.Cell(...).");
        /// <summary>Gets the 1-based row index.</summary>
        public int Row { get; }
        /// <summary>Gets the 1-based column index.</summary>
        public int Column { get; }
        /// <summary>Gets the A1 cell address.</summary>
        public string Address => ExcelA1.CellReference(Row, Column);

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
        /// Sets the cell value, inferring the stored type. Dates and times get a date or duration number format
        /// unless the cell already has one.
        /// </summary>
        public ExcelCell SetValue(object? value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(string value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(double value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(float value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(decimal value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(int value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(long value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(short value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(uint value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(ulong value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(ushort value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(byte value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(sbyte value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(bool value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(DateTime value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(DateTimeOffset value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(TimeSpan value) { Sheet.CellValue(Row, Column, value); return this; }
#if NET6_0_OR_GREATER
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(DateOnly value) { Sheet.CellValue(Row, Column, value); return this; }
        /// <inheritdoc cref="SetValue(object)" />
        public ExcelCell SetValue(TimeOnly value) { Sheet.CellValue(Row, Column, value); return this; }
#endif
        /// <summary>Sets the cell value from a nullable value; <c>null</c> writes an empty value.</summary>
        public ExcelCell SetValue<T>(T? value) where T : struct { Sheet.CellValue(Row, Column, value); return this; }

        /// <summary>Writes a native Excel error value, such as <c>#N/A</c> or <c>#DIV/0!</c>.</summary>
        public ExcelCell SetError(string errorText) { Sheet.CellError(Row, Column, errorText); return this; }

        /// <summary>
        /// Sets the cell formula while keeping any existing value as its cached result. This is not a redaction
        /// API: rebuild the workbook from approved data to remove sensitive historic content.
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
        /// Sets a positive physical font size.
        /// </summary>
        public ExcelCell SetFontSize(OfficeLength fontSize) {
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

        /// <summary>Sets the horizontal alignment.</summary>
        public ExcelCell SetHorizontalAlignment(ExcelHorizontalAlignment alignment) {
            Sheet.CellAlign(Row, Column, alignment);
            return this;
        }

        /// <summary>Sets the vertical alignment.</summary>
        public ExcelCell SetVerticalAlignment(ExcelVerticalAlignment alignment) {
            Sheet.CellVerticalAlign(Row, Column, alignment);
            return this;
        }

        /// <summary>Sets or clears text wrapping.</summary>
        public ExcelCell SetWrapText(bool wrapText = true) {
            Sheet.CellWrapText(Row, Column, wrapText);
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
        /// Sets the solid fill color using a hex value, <c>#RRGGBB</c> or <c>#AARRGGBB</c>.
        /// </summary>
        public ExcelCell SetFillColor(string hexColor) {
            Sheet.CellBackground(Row, Column, hexColor);
            return this;
        }

        /// <summary>Sets the solid fill color. Theme colors stay linked to the workbook theme.</summary>
        public ExcelCell SetFillColor(OfficeColor color) {
            Sheet.CellBackground(Row, Column, color);
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

        /// <summary>Sets selected border edges while preserving other edges and cell styles.</summary>
        public ExcelCell SetBorder(ExcelCellBorders borders, ExcelBorderStyle style, string? hexColor = null) {
            Sheet.CellBorder(Row, Column, borders, style, hexColor);
            return this;
        }

        /// <summary>
        /// Applies diagonal borders to the cell.
        /// </summary>
        public ExcelCell SetDiagonalBorder(ExcelBorderStyle style, string? hexColor = null, bool diagonalUp = true, bool diagonalDown = true) {
            Sheet.CellDiagonalBorder(Row, Column, style, hexColor, diagonalUp, diagonalDown);
            return this;
        }

        /// <summary>
        /// Replaces the cell contents with rich inline text runs.
        /// </summary>
        public ExcelCell SetRichText(params ExcelRichTextRun[] runs) => SetRichText((IEnumerable<ExcelRichTextRun>)runs);

        /// <summary>
        /// Replaces the cell contents with rich inline text runs.
        /// </summary>
        public ExcelCell SetRichText(IEnumerable<ExcelRichTextRun> runs) {
            Sheet.SetRichText(Row, Column, runs);
            return this;
        }

        /// <summary>
        /// Gets rich inline text runs from the cell.
        /// </summary>
        public IReadOnlyList<ExcelRichTextRun> GetRichText() => Sheet.GetRichText(Row, Column);

        /// <summary>Returns whether both handles address the same cell of the same worksheet object.</summary>
        public bool Equals(ExcelCell other) => ReferenceEquals(_sheet, other._sheet) && Row == other.Row && Column == other.Column;

        /// <inheritdoc />
        public override bool Equals(object? obj) => obj is ExcelCell other && Equals(other);

        /// <inheritdoc />
        public override int GetHashCode() {
            unchecked {
                int hash = _sheet == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_sheet);
                return (hash * 397 ^ Row) * 397 ^ Column;
            }
        }

        /// <summary>Returns whether both handles address the same cell.</summary>
        public static bool operator ==(ExcelCell left, ExcelCell right) => left.Equals(right);

        /// <summary>Returns whether the handles address different cells.</summary>
        public static bool operator !=(ExcelCell left, ExcelCell right) => !left.Equals(right);
    }

    /// <summary>
    /// Lightweight object model wrapper for an A1 range: a cell, a rectangle, whole rows or columns, or a list of
    /// such areas separated by spaces or commas.
    /// </summary>
    /// <remarks>
    /// A range with several areas formats every cell of every area and carries its conditional formatting and data
    /// validation as one rule. Operations that need one rectangle, such as merging, copying, sorting, tables, values
    /// and AutoFilter, throw for it.
    /// </remarks>
    public sealed partial class ExcelRange {
        private readonly (int r1, int c1, int r2, int c2)[]? _areas;

        internal ExcelRange(ExcelSheet sheet, string address) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentNullException(nameof(address));

            (int r1, int c1, int r2, int c2) bounds;
            if (address.IndexOf(' ') < 0 && address.IndexOf(',') < 0) {
                bounds = ToBounds(SpreadsheetRangeReference.Parse(address, SpreadsheetAddressDialect.ExcelA1), address);
                Address = ToRangeAddress(bounds.r1, bounds.c1, bounds.r2, bounds.c2);
            } else {
                var areas = ParseAreas(address);
                bounds = areas[0];
                _areas = areas.Length > 1 ? areas : null;
                Address = string.Join(" ", areas.Select(area => ToRangeAddress(area.r1, area.c1, area.r2, area.c2)));
                foreach (var area in areas) {
                    bounds = (Math.Min(bounds.r1, area.r1), Math.Min(bounds.c1, area.c1), Math.Max(bounds.r2, area.r2), Math.Max(bounds.c2, area.c2));
                }
            }
            FirstRow = bounds.r1;
            FirstColumn = bounds.c1;
            LastRow = bounds.r2;
            LastColumn = bounds.c2;
        }

        internal ExcelRange(ExcelSheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            if (firstRow < 1 || firstRow > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (firstColumn < 1 || firstColumn > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(firstColumn));
            if (lastRow < firstRow || lastRow > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(lastRow));
            if (lastColumn < firstColumn || lastColumn > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(lastColumn));
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            Address = ToRangeAddress(firstRow, firstColumn, lastRow, lastColumn);
        }

        /// <summary>Gets the worksheet that owns the range.</summary>
        public ExcelSheet Sheet { get; }
        /// <summary>Gets the A1 range address; areas are separated by spaces.</summary>
        public string Address { get; }
        /// <summary>Gets the first row in the range (of all areas).</summary>
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
        public ExcelCell FirstCell => _areas == null ? new ExcelCell(Sheet, FirstRow, FirstColumn) : new ExcelCell(Sheet, _areas[0].r1, _areas[0].c1);

        /// <summary>Adds data validation rules to the range.</summary>
        public ExcelRangeDataValidationBuilder Validation() => new ExcelRangeDataValidationBuilder(this);

        /// <summary>Adds conditional formatting rules to the range.</summary>
        public ExcelRangeConditionalFormattingBuilder ConditionalFormat() => new ExcelRangeConditionalFormattingBuilder(this);

        /// <summary>
        /// Returns the AutoFilter on this range, creating it when the range has none. An existing filter on the same
        /// range keeps its criteria; a worksheet filter on another range is replaced. A table on the same range
        /// owns the filter.
        /// </summary>
        public ExcelAutoFilter AutoFilter() {
            SingleArea(nameof(AutoFilter));
            return Sheet.EnsureAutoFilter(Reference);
        }

        /// <summary>
        /// Clears selected data and metadata from the range.
        /// </summary>
        public ExcelRange Clear(ExcelClearOptions options = ExcelClearOptions.All) {
            ForEachArea(area => Sheet.ClearRange(area, options));
            return this;
        }

        /// <summary>
        /// Sorts the range by a 1-based column offset.
        /// </summary>
        public ExcelRange SortByColumn(int columnOffset, bool ascending = true, bool hasHeader = true) {
            Sheet.SortRangeByColumn(SingleArea(nameof(SortByColumn)), columnOffset, ascending, hasHeader);
            return this;
        }

        /// <summary>Transactionally inserts a cell block at this range.</summary>
        public ExcelMutationResult Insert(ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.InsertCells(SingleArea(nameof(Insert)), direction, options, cancellationToken);

        /// <summary>Transactionally deletes this cell block.</summary>
        public ExcelMutationResult Delete(ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.DeleteCells(SingleArea(nameof(Delete)), direction, options, cancellationToken);

        /// <summary>Copies this range to a destination top-left cell.</summary>
        public ExcelMutationResult CopyTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.CopyRange(SingleArea(nameof(CopyTo)), destinationTopLeft, options, cancellationToken);

        /// <summary>Moves this range to a destination top-left cell.</summary>
        public ExcelMutationResult MoveTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.MoveRange(SingleArea(nameof(MoveTo)), destinationTopLeft, options, cancellationToken);

        /// <summary>Copies this range to a destination with rows and columns transposed.</summary>
        public ExcelMutationResult TransposeTo(string destinationTopLeft, ExcelMutationPlanOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
            Sheet.TransposeRange(SingleArea(nameof(TransposeTo)), destinationTopLeft, options, cancellationToken);

        /// <summary>
        /// Merges the range.
        /// </summary>
        public ExcelRange Merge() {
            Sheet.MergeRange(SingleArea(nameof(Merge)));
            return this;
        }

        /// <summary>
        /// Removes merge definitions that overlap the range.
        /// </summary>
        public ExcelRange Unmerge() {
            Sheet.UnmergeRange(SingleArea(nameof(Unmerge)));
            return this;
        }

        /// <summary>
        /// Creates an Excel table over the range.
        /// </summary>
        public ExcelTable CreateTable(string name, bool hasHeader = true, ExcelTableStyle style = ExcelTableStyle.TableStyleMedium2, bool includeAutoFilter = true) {
            string resolvedName = Sheet.AddTableAndGetName(SingleArea(nameof(CreateTable)), hasHeader, name, style, includeAutoFilter);
            return Sheet.Table(resolvedName);
        }

        /// <summary>
        /// Applies a number format to every cell in the range.
        /// </summary>
        public ExcelRange SetNumberFormat(string numberFormat) {
            ForEachArea(area => Sheet.FormatRange(area, numberFormat));
            return this;
        }

        /// <summary>
        /// Applies a fill color to every cell in the range.
        /// </summary>
        public ExcelRange SetFillColor(string hexColor) {
            ForEachArea(area => Sheet.FillRange(area, hexColor));
            return this;
        }

        /// <summary>
        /// Applies a two-color linear gradient fill to every cell in the range.
        /// </summary>
        public ExcelRange SetGradientFill(string fromHexColor, string toHexColor, double degree = 0) {
            ForEachArea(area => Sheet.FillRangeGradient(area, fromHexColor, toHexColor, degree));
            return this;
        }

        /// <summary>Applies a solid fill color to every cell in the range. Theme colors stay linked to the workbook theme.</summary>
        public ExcelRange SetFillColor(OfficeColor color) {
            ForEachCell((row, column) => Sheet.CellBackground(row, column, color));
            return this;
        }

        /// <summary>Sets the horizontal alignment of every cell in the range.</summary>
        public ExcelRange SetHorizontalAlignment(ExcelHorizontalAlignment alignment) {
            ForEachCell((row, column) => Sheet.CellAlign(row, column, alignment));
            return this;
        }

        /// <summary>Sets the vertical alignment of every cell in the range.</summary>
        public ExcelRange SetVerticalAlignment(ExcelVerticalAlignment alignment) {
            ForEachCell((row, column) => Sheet.CellVerticalAlign(row, column, alignment));
            return this;
        }

        /// <summary>Sets or clears text wrapping for every cell in the range.</summary>
        public ExcelRange SetWrapText(bool wrapText = true) {
            ForEachCell((row, column) => Sheet.CellWrapText(row, column, wrapText));
            return this;
        }

        /// <summary>Applies a border style to every edge of every cell in the range.</summary>
        public ExcelRange SetBorder(ExcelBorderStyle style, string? hexColor = null) {
            ForEachCell((row, column) => Sheet.CellBorder(row, column, style, hexColor));
            return this;
        }

        /// <summary>Sets selected border edges of every cell in the range while preserving other edges and styles.</summary>
        public ExcelRange SetBorder(ExcelCellBorders borders, ExcelBorderStyle style, string? hexColor = null) {
            ForEachCell((row, column) => Sheet.CellBorder(row, column, borders, style, hexColor));
            return this;
        }

        /// <summary>Applies diagonal borders to every cell in the range.</summary>
        public ExcelRange SetDiagonalBorder(ExcelBorderStyle style, string? hexColor = null, bool diagonalUp = true, bool diagonalDown = true) {
            ForEachCell((row, column) => Sheet.CellDiagonalBorder(row, column, style, hexColor, diagonalUp, diagonalDown));
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
        /// Applies a positive physical font size to every cell in the range.
        /// </summary>
        public ExcelRange SetFontSize(OfficeLength fontSize) {
            if (fontSize <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(fontSize), "Font size must be positive.");
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

        // One write lock and one style scope per range: indexed style lookups and at most one stylesheet save.
        private void ForEachCell(Action<int, int> apply) {
            Sheet.RangeEdit(() => {
                if (_areas == null) {
                    ApplyArea(FirstRow, FirstColumn, LastRow, LastColumn, apply);
                    return;
                }

                foreach (var area in _areas) ApplyArea(area.r1, area.c1, area.r2, area.c2, apply);
            });
        }

        private static void ApplyArea(int firstRow, int firstColumn, int lastRow, int lastColumn, Action<int, int> apply) {
            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    apply(row, column);
                }
            }
        }

        private void ForEachArea(Action<string> apply) {
            if (_areas == null) {
                apply(Address);
                return;
            }

            foreach (var area in _areas) apply(ToRangeAddress(area.r1, area.c1, area.r2, area.c2));
        }

        /// <summary>The range as a reference list for rules, writing single cells without a ":" end.</summary>
        internal string Reference {
            get {
                if (_areas == null) return CompactReference(FirstRow, FirstColumn, LastRow, LastColumn);
                return string.Join(" ", _areas.Select(area => CompactReference(area.r1, area.c1, area.r2, area.c2)));
            }
        }

        private static string CompactReference(int r1, int c1, int r2, int c2) =>
            r1 == r2 && c1 == c2 ? ExcelA1.CellReference(r1, c1) : ToRangeAddress(r1, c1, r2, c2);

        internal string SingleArea(string operation) {
            if (_areas != null) {
                throw new InvalidOperationException($"{operation} needs one rectangular range; '{Address}' has {_areas.Length} areas.");
            }

            return Address;
        }

        private static (int r1, int c1, int r2, int c2)[] ParseAreas(string address) {
            var sequence = SpreadsheetReferenceSequence.Parse(address, SpreadsheetAddressDialect.ExcelA1);
            var areas = new (int r1, int c1, int r2, int c2)[sequence.References.Count];
            for (int index = 0; index < areas.Length; index++) {
                areas[index] = ToBounds(sequence.References[index], address);
            }

            return areas;
        }

        private static (int r1, int c1, int r2, int c2) ToBounds(SpreadsheetRangeReference syntax, string address) {
            if (syntax.Start.SheetName != null || syntax.End?.SheetName != null)
                throw new ArgumentException("Range must not qualify a different worksheet.", nameof(address));
            var end = syntax.End ?? syntax.Start;
            int r1, c1, r2, c2;
            if (syntax.Start.IsCell && end.IsCell) {
                r1 = checked((int)syntax.Start.Row!.Value); c1 = syntax.Start.Column!.Value;
                r2 = checked((int)end.Row!.Value); c2 = end.Column!.Value;
            } else if (syntax.Start.Column.HasValue && end.Column.HasValue && !syntax.Start.Row.HasValue && !end.Row.HasValue) {
                r1 = 1; c1 = syntax.Start.Column.Value; r2 = ExcelA1.MaxRows; c2 = end.Column.Value;
            } else if (syntax.Start.Row.HasValue && end.Row.HasValue && !syntax.Start.Column.HasValue && !end.Column.HasValue) {
                r1 = checked((int)syntax.Start.Row.Value); c1 = 1; r2 = checked((int)end.Row.Value); c2 = ExcelA1.MaxColumns;
            } else {
                throw new ArgumentException("Range must be a cell, a rectangle, whole rows or whole columns.", nameof(address));
            }

            if (r1 > r2) (r1, r2) = (r2, r1);
            if (c1 > c2) (c1, c2) = (c2, c1);
            if (r1 < 1 || r2 > ExcelA1.MaxRows || c1 < 1 || c2 > ExcelA1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(address), $"Range '{address}' is outside the worksheet.");
            return (r1, c1, r2, c2);
        }

        private static string ToRangeAddress(int r1, int c1, int r2, int c2) {
            string start = ExcelA1.CellReference(r1, c1);
            string end = ExcelA1.CellReference(r2, c2);
            return $"{start}:{end}";
        }
    }

    /// <summary>
    /// Lightweight object model wrapper for an Excel table.
    /// </summary>
    public sealed partial class ExcelTable {
        internal ExcelTable(ExcelSheet sheet, string nameOrRange) {
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            NameOrRange = string.IsNullOrWhiteSpace(nameOrRange) ? throw new ArgumentNullException(nameof(nameOrRange)) : nameOrRange;
        }

        /// <summary>Gets the worksheet that owns the table.</summary>
        public ExcelSheet Sheet { get; }
        /// <summary>Gets the table name, display name, or A1 range used to locate the table.</summary>
        public string NameOrRange { get; }
        /// <summary>Gets the table range when it can be resolved.</summary>
        public string? Range => Sheet.GetTableRange(NameOrRange) ?? (ExcelA1.TryParseRange(NameOrRange, out _, out _, out _, out _) ? NameOrRange : null);

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
        private OfficeLength? _fontSize;

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
                if (!value) {
                    _underlineStyle = ExcelUnderlineStyle.None;
                } else if (!_underlineStyle.HasValue || _underlineStyle == ExcelUnderlineStyle.None) {
                    _underlineStyle = ExcelUnderlineStyle.Single;
                }
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
                if (value.HasValue) {
                    _underline = value != ExcelUnderlineStyle.None;
                    UnderlineSpecified = true;
                }
            }
        }
        /// <summary>Gets or sets the run font color as a hex value.</summary>
        public string? FontColor { get; set; }
        /// <summary>Gets or sets the run font name.</summary>
        public string? FontName { get; set; }
        /// <summary>Gets or sets the positive physical run font size; null inherits the cell font size.</summary>
        public OfficeLength? FontSize {
            get => _fontSize;
            set {
                if (value.HasValue && value.Value <= OfficeLength.Zero)
                    throw new ArgumentOutOfRangeException(nameof(FontSize), "Font size must be positive.");
                _fontSize = value;
            }
        }
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
                FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value is double points ? OfficeLength.Points(points) : (OfficeLength?)null,
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
#pragma warning restore RS0026, RS0027
