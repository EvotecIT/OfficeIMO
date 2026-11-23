namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool IsArrayShapingFunction(string function) =>
            function is "TRANSPOSE" or "MMULT" or "SORTBY" or "VSTACK" or "HSTACK" or "TEXTSPLIT";

        private bool TryEvaluateArrayShaping(string function, string arguments, int depth, out FormulaArrayValue array) {
            array = null!;
            // Positions matter for these calls, so empty arguments are kept in place.
            IReadOnlyList<string> args = SplitFormulaArguments(arguments, preserveEmpty: true);
            if (args.Count == 0) return false;
            if (function == "TEXTSPLIT") return TryEvaluateTextSplit(args, out array);
            var inputs = new FormulaArrayValue[args.Count];
            int cells = 0;
            int count = function == "SORTBY" ? 1 : args.Count;
            for (int index = 0; index < count; index++) {
                if (string.IsNullOrWhiteSpace(args[index]) || !TryEvaluateArrayValue(args[index], depth + 1, out inputs[index])) return false;
                cells += inputs[index].Values.Length;
                if (cells > MaxResolvedFormulaRangeCells) return false;
            }
            switch (function) {
                case "TRANSPOSE":
                    return args.Count == 1 && TryTransposeArray(inputs[0], out array);
                case "MMULT":
                    return args.Count == 2 && TryMultiplyArrays(inputs[0], inputs[1], out array);
                case "SORTBY":
                    return TryEvaluateSortBy(args, inputs[0], depth, out array);
                default:
                    return TryStackArrays(inputs, function == "VSTACK", out array);
            }
        }

        private static bool TryTransposeArray(FormulaArrayValue input, out FormulaArrayValue array) {
            var values = new FormulaArgumentValue[input.Values.Length];
            for (int row = 0; row < input.Rows; row++)
                for (int column = 0; column < input.Columns; column++)
                    values[column * input.Rows + row] = NormalizeArrayOutput(input.Values[row * input.Columns + column]);
            array = new FormulaArrayValue(input.Columns, input.Rows, values);
            return true;
        }

        // Every element must be a number; text, logical values and blanks make the product #VALUE!.
        private static bool TryMultiplyArrays(FormulaArrayValue left, FormulaArrayValue right, out FormulaArrayValue array) {
            array = null!;
            foreach (FormulaArgumentValue value in left.Values.Concat(right.Values)) {
                if (value.IsUnresolvedFormula || value.IsUnevaluatedFormulaCache || value.IsError) return false;
                if (!value.HasValue || value.IsBoolean || !value.Number.HasValue) {
                    array = ArrayError("#VALUE!");
                    return true;
                }
                if (double.IsNaN(value.Number.Value) || double.IsInfinity(value.Number.Value)) return false;
            }
            if (left.Columns != right.Rows) {
                array = ArrayError("#VALUE!");
                return true;
            }
            if ((long)left.Rows * right.Columns > MaxResolvedFormulaRangeCells) return false;
            var values = new FormulaArgumentValue[left.Rows * right.Columns];
            for (int row = 0; row < left.Rows; row++) {
                for (int column = 0; column < right.Columns; column++) {
                    double sum = 0;
                    for (int k = 0; k < left.Columns; k++)
                        sum += left.Values[row * left.Columns + k].Number!.Value * right.Values[k * right.Columns + column].Number!.Value;
                    values[row * right.Columns + column] = double.IsInfinity(sum) || double.IsNaN(sum)
                        ? FormulaArgumentValue.Error("#NUM!")
                        : new FormulaArgumentValue(sum, null);
                }
            }
            array = new FormulaArrayValue(left.Rows, right.Columns, values);
            return true;
        }

        /// <summary>
        /// SORTBY with one or more single-row or single-column keys. Keys use SORT's qualified collation: numbers, then
        /// ASCII letter-and-digit text, then logical values, with blanks last in either direction; ties keep source order.
        /// </summary>
        private bool TryEvaluateSortBy(IReadOnlyList<string> args, FormulaArrayValue input, int depth, out FormulaArrayValue array) {
            array = null!;
            if (args.Count < 2) return false;
            if (_formulaEvaluationDepthFrames?.Count > 0
                && _formulaEvaluationDepthFrames.Peek().UsedUnevaluatedFormulaCache) return false;
            var keys = new List<(FormulaArrayValue Values, ArraySortKeyKind[] Kinds, int Order)>();
            bool? byColumns = null;
            // The input and every key share one cell budget, so many large keys can't multiply the memory a call holds.
            long cells = input.Values.Length;
            for (int index = 1; index < args.Count; index += 2) {
                if (string.IsNullOrWhiteSpace(args[index]) || !TryEvaluateArrayValue(args[index], depth + 1, out FormulaArrayValue key)) return false;
                cells += key.Values.Length;
                if (cells > MaxResolvedFormulaRangeCells) return false;
                double order = 1;
                if (index + 1 < args.Count && !TryArrayNumber(args, index + 1, 1, out order)) return false;
                order = Math.Truncate(order);
                bool keyByColumns = key.Rows == 1 && key.Columns > 1;
                int length = keyByColumns ? key.Columns : key.Rows;
                if ((key.Rows != 1 && key.Columns != 1) || (byColumns.HasValue && byColumns.Value != keyByColumns)
                    || length != (keyByColumns ? input.Columns : input.Rows) || (order != 1 && order != -1)) {
                    array = ArrayError("#VALUE!");
                    return true;
                }
                byColumns = keyByColumns;
                var kinds = new ArraySortKeyKind[length];
                for (int i = 0; i < length; i++) {
                    if (!TryGetArraySortKeyKind(key.Values[i], out kinds[i])) return false;
                }
                keys.Add((key, kinds, (int)order));
            }
            int vectors = byColumns == true ? input.Columns : input.Rows;
            int[] selected = Enumerable.Range(0, vectors).ToArray();
            System.Array.Sort(selected, (left, right) => {
                foreach (var key in keys) {
                    int comparison = CompareSortKeys(key.Values.Values[left], key.Kinds[left], key.Values.Values[right], key.Kinds[right], key.Order);
                    if (comparison != 0) return comparison;
                }
                return left.CompareTo(right);
            });
            array = SelectArrayVectors(input, selected, byColumns == true);
            return true;
        }

        private static int CompareSortKeys(FormulaArgumentValue leftValue, ArraySortKeyKind leftKind,
            FormulaArgumentValue rightValue, ArraySortKeyKind rightKind, int order) {
            if (leftKind == ArraySortKeyKind.Blank || rightKind == ArraySortKeyKind.Blank)
                return leftKind == rightKind ? 0 : leftKind == ArraySortKeyKind.Blank ? 1 : -1;
            int comparison = ((int)leftKind).CompareTo((int)rightKind);
            if (comparison == 0) {
                comparison = leftKind == ArraySortKeyKind.Text
                    ? string.Compare(leftValue.Text, rightValue.Text, StringComparison.OrdinalIgnoreCase)
                    : leftValue.Number!.Value.CompareTo(rightValue.Number!.Value);
            }
            return Math.Sign(comparison) * order;
        }

        // VSTACK appends rows and HSTACK appends columns; shorter inputs are padded with #N/A.
        private static bool TryStackArrays(FormulaArrayValue[] inputs, bool vertical, out FormulaArrayValue array) {
            int rows = 0, columns = 0;
            foreach (FormulaArrayValue input in inputs) {
                if (vertical) { rows += input.Rows; columns = Math.Max(columns, input.Columns); }
                else { columns += input.Columns; rows = Math.Max(rows, input.Rows); }
            }
            if ((long)rows * columns > MaxResolvedFormulaRangeCells) { array = null!; return false; }
            var values = new FormulaArgumentValue[rows * columns];
            FormulaArgumentValue missing = FormulaArgumentValue.Error("#N/A");
            for (int i = 0; i < values.Length; i++) values[i] = missing;
            int offset = 0;
            foreach (FormulaArrayValue input in inputs) {
                for (int row = 0; row < input.Rows; row++)
                    for (int column = 0; column < input.Columns; column++) {
                        int targetRow = vertical ? offset + row : row, targetColumn = vertical ? column : offset + column;
                        values[targetRow * columns + targetColumn] = NormalizeArrayOutput(input.Values[row * input.Columns + column]);
                    }
                offset += vertical ? input.Rows : input.Columns;
            }
            array = new FormulaArrayValue(rows, columns, values);
            return true;
        }

        /// <summary>
        /// TEXTSPLIT(text, column delimiter, [row delimiter], [ignore empty], [match mode], [pad with]). Delimiters are
        /// text or array constants of text; at each position the earliest, then longest, delimiter wins.
        /// </summary>
        private bool TryEvaluateTextSplit(IReadOnlyList<string> args, out FormulaArrayValue array) {
            array = null!;
            if (args.Count < 2 || args.Count > 6 || string.IsNullOrWhiteSpace(args[0])) return false;
            if (!TryResolveFormulaValueArgument(args[0], out FormulaArgumentValue source)) return false;
            if (source.IsError) { array = ArrayError(source.ErrorCode!); return true; }
            if (!TryGetTextSplitDelimiters(args, 1, out string[]? columnDelimiters, out FormulaArgumentValue error)
                || !TryGetTextSplitDelimiters(args, 2, out string[]? rowDelimiters, out FormulaArgumentValue rowError)) return false;
            if (!error.IsError) error = rowError;
            if (!TryArrayFlag(args, 3, out bool ignoreEmpty) || !TryArrayNumber(args, 4, 0, out double mode)) return false;
            FormulaArgumentValue padding = FormulaArgumentValue.Error("#N/A");
            if (args.Count > 5 && !string.IsNullOrWhiteSpace(args[5])) {
                if (!TryResolveFormulaValueArgument(args[5], out padding)) return false;
                padding = NormalizeArrayOutput(padding);
            }
            if (error.IsError) { array = ArrayError(error.ErrorCode!); return true; }
            if ((columnDelimiters == null && rowDelimiters == null) || (mode != 0 && mode != 1)) {
                array = ArrayError("#VALUE!");
                return true;
            }
            StringComparison comparison = mode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            string text = FormulaValueToText(source);
            List<string> lines = SplitText(text, rowDelimiters, comparison, ignoreEmpty);
            var rows = new List<List<string>>(lines.Count);
            int columns = 0;
            foreach (string line in lines) {
                List<string> cells = SplitText(line, columnDelimiters, comparison, ignoreEmpty);
                rows.Add(cells);
                columns = Math.Max(columns, cells.Count);
            }
            if (rows.Count == 0 || columns == 0) {
                array = ArrayError("#CALC!");
                return true;
            }
            if ((long)rows.Count * columns > MaxResolvedFormulaRangeCells) return false;
            var values = new FormulaArgumentValue[rows.Count * columns];
            for (int row = 0; row < rows.Count; row++)
                for (int column = 0; column < columns; column++)
                    values[row * columns + column] = column < rows[row].Count ? new FormulaArgumentValue(null, rows[row][column]) : padding;
            array = new FormulaArrayValue(rows.Count, columns, values);
            return true;
        }

        private bool TryGetTextSplitDelimiters(IReadOnlyList<string> args, int index, out string[]? delimiters, out FormulaArgumentValue error) {
            delimiters = null;
            error = default;
            if (index >= args.Count || string.IsNullOrWhiteSpace(args[index])) return true;
            if (!TryEvaluateArrayValue(args[index], 1, out FormulaArrayValue values)) return false;
            var list = new List<string>(values.Values.Length);
            foreach (FormulaArgumentValue value in values.Values) {
                if (value.IsError) { error = value; return true; }
                if (value.IsUnresolvedFormula || !value.HasValue) return false;
                string delimiter = FormulaValueToText(value);
                if (delimiter.Length == 0) { error = FormulaArgumentValue.Error("#VALUE!"); return true; }
                list.Add(delimiter);
            }
            delimiters = list.ToArray();
            return true;
        }

        private static List<string> SplitText(string text, string[]? delimiters, StringComparison comparison, bool ignoreEmpty) {
            var parts = new List<string>();
            if (delimiters == null) {
                if (!ignoreEmpty || text.Length > 0) parts.Add(text);
                return parts;
            }
            int start = 0;
            while (true) {
                int found = -1, length = 0;
                foreach (string delimiter in delimiters) {
                    int position = text.IndexOf(delimiter, start, comparison);
                    if (position >= 0 && (found < 0 || position < found || (position == found && delimiter.Length > length))) {
                        found = position;
                        length = delimiter.Length;
                    }
                }
                string part = found < 0 ? text.Substring(start) : text.Substring(start, found - start);
                if (!ignoreEmpty || part.Length > 0) parts.Add(part);
                if (found < 0) break;
                start = found + length;
            }
            return parts;
        }
    }
}
