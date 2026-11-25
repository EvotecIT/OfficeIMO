using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Reference-backed sources resolve only visited cells. In particular a
        // binary lookup must not materialize its entire input or return range.
        // A source reads a value array, a reference window or one value repeated over its whole shape;
        // rows, columns and slices narrow that window without copying it.
        private sealed class FormulaLookupSource : IReadOnlyList<FormulaArgumentValue> {
            private readonly FormulaArgumentValue[]? _values;
            private readonly int _offset;
            private readonly int _stride;
            private readonly ExcelSheet? _owner;
            private readonly ExcelSheet? _sheet;
            private readonly int _top;
            private readonly int _left;
            private readonly bool _constant;
            // Most sources carry no error, so one is kept boxed rather than widening every source.
            private readonly object? _error;

            private FormulaLookupSource(int rows, int columns, FormulaArgumentValue[]? values, int offset, int stride,
                ExcelSheet? owner, ExcelSheet? sheet, int top, int left, bool constant, FormulaArgumentValue error) {
                Rows = rows;
                Columns = columns;
                _values = values;
                _offset = offset;
                _stride = stride;
                _owner = owner;
                _sheet = sheet;
                _top = top;
                _left = left;
                _constant = constant;
                _error = error.IsError ? error : null;
            }

            /// <summary>Reads row-major values; <paramref name="stride"/> is the width of the stored rows.</summary>
            internal static FormulaLookupSource FromValues(int rows, int columns, FormulaArgumentValue[] values, int stride,
                FormulaArgumentValue error = default) =>
                new(rows, columns, values, 0, stride, null, null, 0, 0, false, error);

            /// <summary>Resolves cells of <paramref name="sheet"/> from <paramref name="top"/>, <paramref name="left"/> as the owner's dependencies.</summary>
            internal static FormulaLookupSource FromReference(ExcelSheet owner, ExcelSheet sheet, int rows, int columns, int top, int left) =>
                new(rows, columns, null, 0, 0, owner, sheet, top, left, false, default);

            internal static FormulaLookupSource FromConstant(FormulaArgumentValue value, FormulaArgumentValue error) =>
                new(1, 1, new[] { value }, 0, 0, null, null, 0, 0, true, error);

            internal int Rows { get; }
            internal int Columns { get; }
            internal FormulaArgumentValue Error => _error is FormulaArgumentValue error ? error : default;
            public int Count => Rows * Columns;
            public FormulaArgumentValue this[int index] => At(index / Columns, index % Columns);

            internal FormulaArgumentValue At(int row, int column) {
                if (_values != null) return _values[_constant ? 0 : _offset + row * _stride + column];
                return _owner!.ResolveFormulaDependency(_sheet!, _top + row, _left + column);
            }

            internal FormulaLookupSource Column(int column) => Slice(0, column, Rows, 1);
            internal FormulaLookupSource Row(int row) => Slice(row, 0, 1, Columns);
            internal FormulaLookupSource Slice(int row, int column, int rows, int columns) {
                if (_owner != null)
                    return new(rows, columns, null, 0, 0, _owner, _sheet, _top + row, _left + column, false, default);
                if (_constant) return new(rows, columns, _values, 0, 0, null, null, 0, 0, true, default);
                return new(rows, columns, _values, _offset + row * _stride + column, _stride, null, null, 0, 0, false, default);
            }

            public IEnumerator<FormulaArgumentValue> GetEnumerator() {
                for (int index = 0; index < Count; index++) yield return this[index];
            }
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private bool TryResolveLookupSource(string token, ref int budget, out FormulaLookupSource source,
            bool preserveScalarValue = false) {
            source = null!;
            if (TryResolveFormulaRangeReferenceWithError(token, out ExcelSheet sheet,
                out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue referenceError)) {
                if (!TryGetFormulaRangeCellCount(r1, c1, r2, c2, out int count) || count > budget) return false;
                budget -= count;
                int columns = c2 - c1 + 1;
                if (TryGetPlainCalculatedRange(sheet, r1, c1, r2, c2, out FormulaArgumentValue[] plain)) {
                    source = FormulaLookupSource.FromValues(r2 - r1 + 1, columns, plain, columns);
                    return true;
                }
                source = FormulaLookupSource.FromReference(this, sheet, r2 - r1 + 1, columns, r1, c1);
                return true;
            }
            if (referenceError.IsError) {
                source = FormulaLookupSource.FromConstant(referenceError, referenceError);
                return true;
            }
            if (!TryEvaluateArrayValue(token, 0, out FormulaArrayValue array, preserveScalarValue)
                || array.Values.Length > budget) return false;
            budget -= array.Values.Length;
            FormulaArgumentValue error = array.Values.Length == 1 && array.Values[0].IsError
                && !token.TrimStart().StartsWith("{", StringComparison.Ordinal) ? array.Values[0] : default;
            source = FormulaLookupSource.FromValues(array.Rows, array.Columns, array.Values, array.Columns, error);
            return true;
        }

        /// <summary>
        /// Returns a lookup range of plain values read once for the pass. Reading such a range ahead evaluates nothing
        /// and charges no dependency depth, so it matches the lazy read; ranges with formulas or array output stay lazy.
        /// </summary>
        private bool TryGetPlainCalculatedRange(ExcelSheet sheet, int r1, int c1, int r2, int c2, out FormulaArgumentValue[] values) {
            values = null!;
            if (_formulaCalculationContext is not { IsDisposed: false } context || _formulaEvaluationCache == null) return false;
            var key = ((object)sheet._worksheetPartValue, r1, c1, r2, c2);
            if (context.Ranges.TryGetValue(key, out var range)) {
                if (range.Depth != 0) return false;
                values = range.Values;
                return true;
            }
            if (context.LazyRanges.Contains(key)) return false;
            if (_excelDocument.WorkbookPartRoot.CellMetadataPart != null
                || sheet._formulaEvaluationCache == null
                || sheet.GetFixedArraySheetIndex().Owners.Count != 0) {
                context.LazyRanges.Add(key);
                return false;
            }

            var resolved = new FormulaArgumentValue[(r2 - r1 + 1) * (c2 - c1 + 1)];
            int index = 0;
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    Cell? cell = sheet.TryGetDomCell(row, column);
                    if (cell != null && ExcelWorksheetScan.FindFormula(cell) != null) {
                        context.LazyRanges.Add(key);
                        return false;
                    }
                    FormulaArgumentValue value = sheet.ResolveCellArgument(row, column);
                    if (value.IsUnresolvedFormula || value.IsUnevaluatedFormulaCache) {
                        context.LazyRanges.Add(key);
                        return false;
                    }
                    resolved[index++] = value;
                }
            }
            context.Ranges[key] = (resolved, 0);
            values = resolved;
            return true;
        }

        private static FormulaLookupSource LookupScalarSource(FormulaArgumentValue value) =>
            FormulaLookupSource.FromConstant(value, value.IsError ? value : default);

        private static bool TryMaterializeLookupArray(FormulaLookupSource source, out FormulaArrayValue array) {
            array = null!;
            var values = new FormulaArgumentValue[source.Count];
            for (int index = 0; index < values.Length; index++) {
                // Range consumers retain empty references; output-cell publication applies zero coercion.
                values[index] = source[index];
                if (values[index].IsUnresolvedFormula) return false;
            }
            array = new FormulaArrayValue(source.Rows, source.Columns, values);
            return true;
        }

        private static bool LookupValuesEqual(FormulaArgumentValue left, FormulaArgumentValue right) {
            if (left.IsError || right.IsError || left.IsUnresolvedFormula || right.IsUnresolvedFormula) return false;
            if (left.HasValue != right.HasValue || left.IsBoolean != right.IsBoolean || left.Number.HasValue != right.Number.HasValue) return false;
            if (left.Number.HasValue) return left.Number.Value == right.Number!.Value;
            return string.Equals(left.Text ?? string.Empty, right.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
    }
}
