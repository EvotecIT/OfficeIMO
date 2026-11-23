using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// LOOKUP in its vector and array forms. Like VLOOKUP's approximate mode it returns the last value not greater
        /// than the lookup value in a sorted vector, so unsorted input gives the same answer VLOOKUP would.
        /// </summary>
        private bool TryEvaluateLookupFunctionValue(string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 2 || tokens.Count > 3 || tokens.Any(string.IsNullOrWhiteSpace)
                || !TryResolveLookupArgument(tokens[0], out FormulaArgumentValue lookupValue)
                || lookupValue.IsUnresolvedFormula) return false;
            int budget = MaxResolvedFormulaRangeCells;
            if (!TryResolveLookupSource(tokens[1], ref budget, out FormulaLookupSource source)) return false;
            FormulaLookupSource? results = null;
            if (tokens.Count == 3 && !TryResolveLookupSource(tokens[2], ref budget, out results)) return false;
            if (_formulaReferenceError.IsError) return true;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (source.Error.IsError) { result = source.Error; return true; }
            if (results?.Error.IsError == true) { result = results.Error; return true; }

            FormulaLookupSource searched;
            if (results != null) {
                // The vector form searches a single row or column and reads the same position from the result vector.
                if ((source.Rows != 1 && source.Columns != 1) || (results.Rows != 1 && results.Columns != 1)) return false;
                searched = source;
            } else {
                // The array form searches the first row of a wide array or the first column otherwise, and returns
                // the matching entry of the last row or column.
                searched = source.Columns > source.Rows ? source.Row(0) : source.Column(0);
            }

            lookupValue = NormalizeLookupResult(lookupValue);
            if (!TryFindLookupPosition(lookupValue, searched, -1, 1, out int position, out FormulaArgumentValue failure,
                sortedApproximate: true)) {
                if (failure.IsUnresolvedFormula) return false;
                result = failure.IsError ? failure : FormulaArgumentValue.Error("#N/A");
                return true;
            }

            if (results != null) {
                if (position > results.Count) return false;
                result = results[position - 1].AsScalarValue();
            } else {
                result = (source.Columns > source.Rows
                    ? source.At(source.Rows - 1, position - 1)
                    : source.At(position - 1, source.Columns - 1)).AsScalarValue();
            }
            return !result.IsUnresolvedFormula;
        }

        /// <summary>ADDRESS: the text of a cell reference in A1 or R1C1 style, optionally qualified by a sheet name.</summary>
        private bool TryEvaluateAddressValue(string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 2 || tokens.Count > 5
                || string.IsNullOrWhiteSpace(tokens[0]) || string.IsNullOrWhiteSpace(tokens[1])) return false;
            double[] numbers = { 0, 0, 1, 1 };
            FormulaArgumentValue firstError = default;
            for (int index = 0; index < Math.Min(tokens.Count, 4); index++) {
                if (string.IsNullOrWhiteSpace(tokens[index])) continue;
                if (!TryResolveFormulaValueArgument(tokens[index], out FormulaArgumentValue value)) return false;
                if (value.IsError) {
                    if (!firstError.IsError) firstError = value;
                    continue;
                }
                if (!IsFiniteFormulaNumber(value)) return false;
                numbers[index] = value.Number!.Value;
            }
            string? sheetText = null;
            if (tokens.Count == 5 && !string.IsNullOrWhiteSpace(tokens[4])) {
                if (!TryResolveFormulaValueArgument(tokens[4], out FormulaArgumentValue sheet)) return false;
                if (sheet.IsError) {
                    if (!firstError.IsError) firstError = sheet;
                } else {
                    sheetText = FormulaValueToText(sheet);
                }
            }
            if (firstError.IsError) { result = firstError; return true; }

            double row = Math.Truncate(numbers[0]), column = Math.Truncate(numbers[1]), kind = Math.Truncate(numbers[2]);
            bool a1 = numbers[3] != 0;
            if (row < 1 || row > ExcelA1.MaxRows || column < 1 || column > ExcelA1.MaxColumns || kind < 1 || kind > 4) {
                result = FormulaArgumentValue.Error("#VALUE!");
                return true;
            }
            bool absoluteRow = kind is 1 or 2, absoluteColumn = kind is 1 or 3;
            int r = (int)row, c = (int)column;
            var text = new StringBuilder();
            if (sheetText != null) text.Append(FormatAddressSheetName(sheetText)).Append('!');
            if (a1) {
                if (absoluteColumn) text.Append('$');
                text.Append(ExcelA1.ColumnIndexToLetters(c));
                if (absoluteRow) text.Append('$');
                text.Append(r.ToString(CultureInfo.InvariantCulture));
            } else {
                text.Append('R').Append(absoluteRow ? r.ToString(CultureInfo.InvariantCulture) : "[" + r.ToString(CultureInfo.InvariantCulture) + "]");
                text.Append('C').Append(absoluteColumn ? c.ToString(CultureInfo.InvariantCulture) : "[" + c.ToString(CultureInfo.InvariantCulture) + "]");
            }
            result = new FormulaArgumentValue(null, text.ToString());
            return true;
        }

        // A name of letters, digits, underscores and periods that does not start with a digit needs no quotes.
        private static string FormatAddressSheetName(string name) {
            if (name.Length == 0) return name;
            bool plain = !char.IsDigit(name[0]);
            foreach (char character in name) {
                if (!char.IsLetterOrDigit(character) && character != '_' && character != '.') {
                    plain = false;
                    break;
                }
            }
            if (plain && (ExcelA1.TryParseCellReferenceFast(name.ToUpperInvariant(), out _, out _) || LooksLikeR1C1(name))) plain = false;
            return plain ? name : "'" + name.Replace("'", "''") + "'";
        }

        private static bool LooksLikeR1C1(string name) {
            string upper = name.ToUpperInvariant();
            return upper is "R" or "C" || TryParseR1C1Cell(upper, 1, 1, out _, out _);
        }

        /// <summary>INDIRECT as a single-cell value; a multi-cell reference is an argument, not a scalar, in this subset.</summary>
        private bool TryEvaluateIndirectValue(string args, out FormulaArgumentValue result) {
            result = default;
            int? currentRow = _formulaEvaluationCellReference != null
                && TryParseCellReference(_formulaEvaluationCellReference, out int row, out _) ? row : null;
            if (!TryResolveIndirectRange(args, currentRow, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2, out result)) return false;
            if (result.IsError) return true;
            if (r1 != r2 || c1 != c2) return false;
            result = ResolveFormulaDependency(sheet, r1, c1);
            return !result.IsUnresolvedFormula;
        }

        /// <summary>
        /// Resolves the reference INDIRECT names. Text that cannot be a reference is #REF!, as is a qualified
        /// reference to a sheet the workbook lacks; text this subset cannot read, such as whole columns, is unsupported.
        /// </summary>
        private bool TryResolveIndirectRange(string args, int? currentRow, out ExcelSheet sheet,
            out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue error) {
            sheet = this;
            r1 = c1 = r2 = c2 = 0;
            error = default;
            if (_offsetFormulaReferenceDepth >= 32 || !HasSufficientFormulaExecutionStack()) return false;
            _offsetFormulaReferenceDepth++;
            FormulaArgumentValue previousError = _formulaReferenceError;
            _formulaReferenceError = default;
            try {
                IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
                if (tokens.Count < 1 || tokens.Count > 2 || string.IsNullOrWhiteSpace(tokens[0])) return false;
                if (!TryResolveFormulaValueArgument(tokens[0], out FormulaArgumentValue reference)) return false;
                bool a1 = true;
                if (tokens.Count == 2 && !string.IsNullOrWhiteSpace(tokens[1])) {
                    if (!TryResolveFormulaValueArgument(tokens[1], out FormulaArgumentValue style)) return false;
                    if (style.IsError) { if (!reference.IsError) error = style; }
                    else if (!IsFiniteFormulaNumber(style)) return false;
                    else a1 = style.Number!.Value != 0;
                }
                if (reference.IsError) { error = reference; return true; }
                if (error.IsError) return true;
                if (reference.Number.HasValue || reference.IsBoolean) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }

                string text = (reference.Text ?? string.Empty).Trim();
                if (text.Length == 0 || text.IndexOf('(') >= 0 || text.IndexOf('"') >= 0) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                if (!a1) {
                    if (!TryConvertR1C1Reference(text, out string? converted)) {
                        error = FormulaArgumentValue.Error("#REF!");
                        return true;
                    }
                    text = converted!;
                }
                if (TryResolveFormulaRangeReference(text, currentRow, out sheet, out r1, out c1, out r2, out c2)) return true;
                // A qualified reference whose sheet is missing names nothing.
                FormulaReferenceSyntax syntax = GetFormulaReferenceSyntax(text);
                if (syntax.Split && syntax.SheetName != null && !TryGetFormulaReferenceSheet(syntax.SheetName, out _)
                    && (syntax.IsRange || syntax.CellRow > 0)) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                // Neither a reference nor a name can hold other characters, such as spaces, outside its sheet qualifier.
                string unqualified = text.Substring(text.LastIndexOf('!') + 1);
                foreach (char character in unqualified) {
                    if (!char.IsLetterOrDigit(character) && character is not ('$' or '_' or '.' or ':' or '\\')) {
                        error = FormulaArgumentValue.Error("#REF!");
                        return true;
                    }
                }
                return false;
            } finally {
                _formulaReferenceError = previousError;
                _offsetFormulaReferenceDepth--;
            }
        }

        /// <summary>Rewrites an R1C1 cell or range, optionally sheet-qualified, as A1 relative to the formula cell.</summary>
        private bool TryConvertR1C1Reference(string text, out string? converted) {
            converted = null;
            int currentRow = 1, currentColumn = 1;
            if (_formulaEvaluationCellReference != null) TryParseCellReference(_formulaEvaluationCellReference, out currentRow, out currentColumn);
            string prefix = string.Empty, reference = text;
            int bang = text.LastIndexOf('!');
            if (bang >= 0) {
                prefix = text.Substring(0, bang + 1);
                reference = text.Substring(bang + 1);
            }
            string[] parts = reference.ToUpperInvariant().Split(':');
            if (parts.Length > 2) return false;
            var builder = new StringBuilder(prefix);
            for (int index = 0; index < parts.Length; index++) {
                if (!TryParseR1C1Cell(parts[index], currentRow, currentColumn, out int row, out int column)) return false;
                if (index > 0) builder.Append(':');
                builder.Append(ExcelA1.CellReference(row, column));
            }
            converted = builder.ToString();
            return true;
        }

        // R and C each take an absolute number, a bracketed offset from the formula cell, or nothing for the same row or column.
        private static bool TryParseR1C1Cell(string text, int currentRow, int currentColumn, out int row, out int column) {
            row = column = 0;
            int position = 0;
            if (position >= text.Length || text[position] != 'R') return false;
            position++;
            if (!TryParseR1C1Part(text, ref position, currentRow, ExcelA1.MaxRows, out row)) return false;
            if (position >= text.Length || text[position] != 'C') return false;
            position++;
            if (!TryParseR1C1Part(text, ref position, currentColumn, ExcelA1.MaxColumns, out column)) return false;
            return position == text.Length;
        }

        private static bool TryParseR1C1Part(string text, ref int position, int current, int maximum, out int value) {
            value = current;
            if (position >= text.Length || (text[position] != '[' && !char.IsDigit(text[position]))) return true;
            bool relative = text[position] == '[';
            if (relative) position++;
            int start = position;
            if (relative && position < text.Length && (text[position] == '-' || text[position] == '+')) position++;
            while (position < text.Length && char.IsDigit(text[position])) position++;
            if (!int.TryParse(text.Substring(start, position - start), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int number)) return false;
            if (relative) {
                if (position >= text.Length || text[position] != ']') return false;
                position++;
                value = current + number;
            } else {
                value = number;
            }
            return value >= 1 && value <= maximum;
        }
    }
}
