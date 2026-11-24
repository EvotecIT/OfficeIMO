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
                SplitIndirectQualifier(text, out string qualifier, out string unqualified);
                // A quoted sheet name may hold parentheses; the reference after it cannot.
                if (unqualified.Length == 0 || unqualified.IndexOf('(') >= 0 || unqualified.IndexOf('"') >= 0) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                if (!a1) {
                    // The style names the cell notation; a defined name resolves the same in either style, and the
                    // name lookup checks a qualified name's sheet scope.
                    if (!TryConvertR1C1Reference(qualifier, unqualified, out string? converted)) {
                        if (TryResolveDefinedNameRange(text, currentRow, out sheet, out r1, out c1, out r2, out c2)) return true;
                        if (TryTakeIndirectReferenceError(out error)) return true;
                        // R1C1 cell syntax that converts to nothing lies outside the grid, and A1 cell syntax is the
                        // wrong style; whole rows or columns and unknown names are unsupported, as in A1 style.
                        if (IsR1C1LineSyntax(unqualified)) return false;
                        if (IsR1C1CellSyntax(unqualified) || TryClassifyA1Reference(unqualified, out _)) {
                            error = FormulaArgumentValue.Error("#REF!");
                            return true;
                        }
                        return HasOnlyReferenceCharacters(unqualified, out error);
                    }
                    text = converted!;
                }
                if (TryResolveFormulaRangeReference(text, currentRow, out sheet, out r1, out c1, out r2, out c2)) return true;
                // A name whose own reference is an error, such as an INDIRECT of invalid text, passes that error on.
                if (TryTakeIndirectReferenceError(out error)) return true;
                // A qualified reference whose sheet is missing names nothing.
                FormulaReferenceSyntax syntax = GetFormulaReferenceSyntax(text);
                if (syntax.Split && syntax.SheetName != null && !TryGetFormulaReferenceSheet(syntax.SheetName, out _)
                    && (syntax.IsRange || syntax.CellRow > 0)) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                if (TryClassifyA1Reference(unqualified, out bool outsideGrid) && outsideGrid) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                return HasOnlyReferenceCharacters(unqualified, out error);
            } finally {
                _formulaReferenceError = previousError;
                _offsetFormulaReferenceDepth--;
            }
        }

        private bool TryTakeIndirectReferenceError(out FormulaArgumentValue error) {
            error = _formulaReferenceError;
            return error.IsError;
        }

        /// <summary>
        /// Neither a reference nor a name can hold other characters, such as spaces, outside its sheet qualifier:
        /// such text is #REF!. Anything else is a name or reference this subset can't resolve and stays unsupported.
        /// </summary>
        private static bool HasOnlyReferenceCharacters(string reference, out FormulaArgumentValue error) {
            error = default;
            foreach (char character in reference) {
                if (!char.IsLetterOrDigit(character) && character is not ('$' or '_' or '.' or ':' or '\\')) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
            }
            return false;
        }

        // A whole row (R2, R[1]:R3) or column (C2): R or C alone with an optional number or bracketed offset.
        private static bool IsR1C1LineSyntax(string reference) {
            string[] parts = reference.ToUpperInvariant().Split(':');
            if (parts.Length > 2) return false;
            foreach (string part in parts) {
                int position = 0;
                if (part.Length == 0 || part[0] is not ('R' or 'C')) return false;
                if (!SkipR1C1Marker(part, ref position, part[0]) || position != part.Length) return false;
            }
            return true;
        }

        // R and C each with an optional number or bracketed offset, for one cell or a range of two.
        private static bool IsR1C1CellSyntax(string reference) {
            string[] parts = reference.ToUpperInvariant().Split(':');
            if (parts.Length > 2) return false;
            foreach (string part in parts) {
                int position = 0;
                if (!SkipR1C1Marker(part, ref position, 'R') || !SkipR1C1Marker(part, ref position, 'C') || position != part.Length) return false;
            }
            return true;
        }

        private static bool SkipR1C1Marker(string text, ref int position, char marker) {
            if (position >= text.Length || text[position] != marker) return false;
            position++;
            if (position < text.Length && text[position] == '[') {
                int close = text.IndexOf(']', position);
                if (close < 0) return false;
                position = close + 1;
                return true;
            }
            while (position < text.Length && char.IsDigit(text[position])) position++;
            return true;
        }

        // Whether text is an A1 cell, or a range of two, with column letters up to three long; outside is set when a
        // row or column lies past the grid.
        private static bool TryClassifyA1Reference(string reference, out bool outside) {
            outside = false;
            string[] parts = reference.Split(':');
            if (parts.Length > 2) return false;
            foreach (string part in parts) {
                int position = 0;
                if (position < part.Length && part[position] == '$') position++;
                int letters = position;
                long column = 0;
                while (position < part.Length && part[position] is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z')) {
                    column = column * 26 + (char.ToUpperInvariant(part[position]) - 'A' + 1);
                    position++;
                }
                letters = position - letters;
                if (letters is < 1 or > 3) return false;
                if (position < part.Length && part[position] == '$') position++;
                int digits = position;
                while (position < part.Length && char.IsDigit(part[position])) position++;
                if (position == digits || position != part.Length) return false;
                string rowText = part.Substring(digits);
                bool rowInside = rowText.Length <= 7 && int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out int row)
                    && row >= 1 && row <= ExcelA1.MaxRows;
                if (!rowInside || column > ExcelA1.MaxColumns) outside = true;
            }
            return true;
        }

        /// <summary>
        /// Splits reference text after its sheet qualifier, which keeps its '!'. A quoted qualifier ends at its closing
        /// quote, so a sheet name may contain '!' or parentheses.
        /// </summary>
        private static void SplitIndirectQualifier(string text, out string qualifier, out string reference) {
            int end = -1;
            if (text.StartsWith("'", StringComparison.Ordinal)) {
                for (int index = 1; index < text.Length; index++) {
                    if (text[index] != '\'') continue;
                    if (index + 1 < text.Length && text[index + 1] == '\'') { index++; continue; }
                    end = index + 1 < text.Length && text[index + 1] == '!' ? index + 1 : -1;
                    break;
                }
            } else {
                end = text.LastIndexOf('!');
            }
            qualifier = end >= 0 ? text.Substring(0, end + 1) : string.Empty;
            reference = end >= 0 ? text.Substring(end + 1) : text;
        }

        /// <summary>Rewrites an R1C1 cell or range, optionally sheet-qualified, as A1 relative to the formula cell.</summary>
        private bool TryConvertR1C1Reference(string prefix, string reference, out string? converted) {
            converted = null;
            int currentRow = 1, currentColumn = 1;
            if (_formulaEvaluationCellReference != null) TryParseCellReference(_formulaEvaluationCellReference, out currentRow, out currentColumn);
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
