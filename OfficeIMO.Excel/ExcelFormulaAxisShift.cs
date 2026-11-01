using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Moves the A1 references of plain formulas through a whole-row or whole-column edit without a syntax tree.
    /// </summary>
    /// <remarks>
    /// Only unqualified references in formulas made of functions, numbers and operators are handled, and a reference
    /// a deletion cuts through is left alone. Anything else makes <see cref="TryRewrite"/> return false, and the
    /// caller falls back to the full reference grammar, so this type never decides a case the grammar owns.
    /// </remarks>
    internal readonly struct ExcelFormulaAxisShift {
        [ThreadStatic]
        private static StringBuilder? _builder;

        private readonly bool _columns;
        private readonly int _first;
        private readonly int _lastDeleted;
        private readonly int _delta;

        private ExcelFormulaAxisShift(bool columns, int first, int lastDeleted, int delta) {
            _columns = columns;
            _first = first;
            _lastDeleted = lastDeleted;
            _delta = delta;
        }

        internal static ExcelFormulaAxisShift Rows(int first, int? lastDeleted, int delta) =>
            new ExcelFormulaAxisShift(columns: false, first, lastDeleted ?? 0, delta);

        internal static ExcelFormulaAxisShift Columns(int first, int count, bool deleting) =>
            new ExcelFormulaAxisShift(columns: true, first, deleting ? first + count - 1 : 0, deleting ? -count : count);

        private bool Deleting => _lastDeleted != 0;

        /// <summary>Checks that <paramref name="formula"/> is plain and stays inside the sheet, without rewriting it.</summary>
        internal bool TryInspect(string formula, out bool hasReference) =>
            Scan(formula, rewrite: false, out _, out hasReference);

        internal bool TryRewrite(string formula, out string rewritten) =>
            Scan(formula, rewrite: true, out rewritten, out _);

        private bool Scan(string formula, bool rewrite, out string rewritten, out bool hasReference) {
            rewritten = formula;
            hasReference = false;
            StringBuilder? output = null;
            int copied = 0;
            int index = 0;
            while (index < formula.Length) {
                char value = formula[index];
                if (value >= '0' && value <= '9') {
                    while (index < formula.Length && formula[index] >= '0' && formula[index] <= '9') index++;
                    if (index < formula.Length && !IsDelimiter(formula[index])) return false;
                    continue;
                }
                if (value != '$' && !(value >= 'A' && value <= 'Z')) {
                    if (!IsDelimiter(value)) return false;
                    index++;
                    continue;
                }

                if (!TryReadPoint(formula, ref index, out ReferencePoint start, out bool function)) return false;
                if (function) continue;
                ReferencePoint end = start;
                if (index < formula.Length && formula[index] == ':') {
                    index++;
                    if (!TryReadPoint(formula, ref index, out end, out function) || function) return false;
                }
                // LOG10( reads like a cell, and only the grammar knows which such names are functions.
                if (index < formula.Length && (formula[index] == '(' || !IsDelimiter(formula[index]))) return false;
                hasReference = true;

                int low = _columns ? start.Column : start.Row;
                int high = _columns ? end.Column : end.Row;
                if (low > high) return false;
                if (high < _first) continue;

                int targetLow = low;
                int targetHigh = high + _delta;
                if (Deleting) {
                    if (low <= _lastDeleted) return false;
                    targetLow = low + _delta;
                } else if (low >= _first) {
                    targetLow = low + _delta;
                }
                if (targetHigh > (_columns ? ExcelA1.MaxColumns : ExcelA1.MaxRows)) return false;

                if (!rewrite) continue;
                output ??= AcquireBuilder(formula.Length + 8);
                copied = Append(output, formula, copied, start, targetLow);
                if (end.Offset != start.Offset) copied = Append(output, formula, copied, end, targetHigh);
            }

            if (output == null) return true;
            output.Append(formula, copied, formula.Length - copied);
            rewritten = output.ToString();
            return true;
        }

        private int Append(StringBuilder output, string formula, int copied, ReferencePoint point, int target) {
            int offset = _columns ? point.ColumnOffset : point.RowOffset;
            int length = _columns ? point.ColumnLength : point.RowLength;
            output.Append(formula, copied, offset - copied);
            if (_columns) output.Append(ExcelA1.ColumnIndexToLetters(target));
            else output.Append(target.ToString(CultureInfo.InvariantCulture));
            return offset + length;
        }

        private static StringBuilder AcquireBuilder(int capacity) {
            StringBuilder builder = _builder ??= new StringBuilder(capacity);
            builder.Clear();
            return builder;
        }

        private static bool TryReadPoint(string formula, ref int index, out ReferencePoint point, out bool function) {
            point = default;
            function = false;
            int offset = index;
            bool absoluteColumn = index < formula.Length && formula[index] == '$';
            if (absoluteColumn) index++;
            int columnOffset = index;
            int column = 0;
            while (index < formula.Length && formula[index] >= 'A' && formula[index] <= 'Z') {
                if (index - columnOffset >= 3) {
                    if (absoluteColumn) return false;
                    return ReadFunctionName(formula, ref index, out function);
                }
                column = column * 26 + formula[index++] - 'A' + 1;
            }
            int columnLength = index - columnOffset;
            if (columnLength == 0) return false;
            if (!absoluteColumn && index < formula.Length && formula[index] == '(') {
                function = true;
                return true;
            }

            if (index < formula.Length && formula[index] == '$') index++;
            int rowOffset = index;
            int row = 0;
            while (index < formula.Length && formula[index] >= '0' && formula[index] <= '9') {
                if (index - rowOffset >= 7) return false;
                row = row * 10 + formula[index++] - '0';
            }
            int rowLength = index - rowOffset;
            if (rowLength == 0 || row == 0 || row > ExcelA1.MaxRows || column > ExcelA1.MaxColumns) return false;
            point = new ReferencePoint(offset, columnOffset, columnLength, rowOffset, rowLength, column, row);
            return true;
        }

        private static bool ReadFunctionName(string formula, ref int index, out bool function) {
            while (index < formula.Length && formula[index] >= 'A' && formula[index] <= 'Z') index++;
            function = index < formula.Length && formula[index] == '(';
            return function;
        }

        private static bool IsDelimiter(char value) => value switch {
            '(' or ')' or ',' or '+' or '-' or '*' or '/' or '^' or '&' or '=' or '<' or '>' or '%' => true,
            _ => false
        };

        private readonly struct ReferencePoint {
            internal ReferencePoint(int offset, int columnOffset, int columnLength, int rowOffset, int rowLength, int column, int row) {
                Offset = offset;
                ColumnOffset = columnOffset;
                ColumnLength = columnLength;
                RowOffset = rowOffset;
                RowLength = rowLength;
                Column = column;
                Row = row;
            }

            internal int Offset { get; }
            internal int ColumnOffset { get; }
            internal int ColumnLength { get; }
            internal int RowOffset { get; }
            internal int RowLength { get; }
            internal int Column { get; }
            internal int Row { get; }
        }
    }
}
