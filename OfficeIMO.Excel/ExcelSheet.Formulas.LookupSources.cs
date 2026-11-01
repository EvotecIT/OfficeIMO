using System.Collections;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Reference-backed sources resolve only visited cells. In particular a
        // binary lookup must not materialize its entire input or return range.
        private sealed class FormulaLookupSource : IReadOnlyList<FormulaArgumentValue> {
            private readonly Func<int, int, FormulaArgumentValue> _read;

            internal FormulaLookupSource(int rows, int columns, Func<int, int, FormulaArgumentValue> read,
                FormulaArgumentValue error = default) {
                Rows = rows;
                Columns = columns;
                _read = read;
                Error = error;
            }

            internal int Rows { get; }
            internal int Columns { get; }
            internal FormulaArgumentValue Error { get; }
            public int Count => Rows * Columns;
            public FormulaArgumentValue this[int index] => _read(index / Columns, index % Columns);
            internal FormulaArgumentValue At(int row, int column) => _read(row, column);
            internal FormulaLookupSource Column(int column) => new(Rows, 1, (row, _) => _read(row, column));
            internal FormulaLookupSource Row(int row) => new(1, Columns, (_, column) => _read(row, column));
            internal FormulaLookupSource Slice(int row, int column, int rows, int columns) =>
                new(rows, columns, (r, c) => _read(row + r, column + c));

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
                source = new FormulaLookupSource(r2 - r1 + 1, c2 - c1 + 1,
                    (row, column) => ResolveFormulaDependency(sheet, r1 + row, c1 + column));
                return true;
            }
            if (referenceError.IsError) {
                source = new FormulaLookupSource(1, 1, (_, _) => referenceError, referenceError);
                return true;
            }
            // SUMPRODUCT distinguishes a blank scalar argument from a zero
            // materialized by ordinary array-value conversion.
            if (preserveScalarValue && !IsSupportedArrayExpression(token)) {
                if (budget < 1 || !TryResolveFormulaArgument(token, out FormulaArgumentValue scalar)
                    || scalar.IsUnresolvedFormula) return false;
                budget--;
                source = LookupScalarSource(scalar);
                return true;
            }
            if (!TryEvaluateArrayValue(token, 0, out FormulaArrayValue array) || array.Values.Length > budget) return false;
            budget -= array.Values.Length;
            FormulaArgumentValue error = array.Values.Length == 1 && array.Values[0].IsError
                && !token.TrimStart().StartsWith("{", StringComparison.Ordinal) ? array.Values[0] : default;
            source = new FormulaLookupSource(array.Rows, array.Columns,
                (row, column) => array.Values[row * array.Columns + column], error);
            return true;
        }

        private static FormulaLookupSource LookupScalarSource(FormulaArgumentValue value) =>
            new(1, 1, (_, _) => value, value.IsError ? value : default);

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
