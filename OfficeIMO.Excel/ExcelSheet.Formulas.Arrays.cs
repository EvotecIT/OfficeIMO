using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<Dictionary<string, FormulaArgumentValue>, ArrayCalculationContext> ArrayCalculationContexts = new System.Runtime.CompilerServices.ConditionalWeakTable<Dictionary<string, FormulaArgumentValue>, ArrayCalculationContext>();

        private sealed class ArrayCalculationContext {
            internal readonly Dictionary<string, FixedArraySheetIndex> Sheets = new Dictionary<string, FixedArraySheetIndex>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, FormulaArrayValue> Results = new Dictionary<string, FormulaArrayValue>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, DynamicArrayPlan> DynamicPlans = new Dictionary<string, DynamicArrayPlan>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<long, FormulaArgumentValue> DynamicCells = new Dictionary<long, FormulaArgumentValue>();
            internal readonly HashSet<long> DynamicRetiredCells = new HashSet<long>();
            internal bool PlanningDynamicOwners;
            internal bool DynamicOwnersPlanned;
        }

        private bool TryGetCalculatedArray(Cell cell, out FormulaArrayValue array) {
            array = null!;
            return cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array && _formulaEvaluationCache != null
                && ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache).Results.TryGetValue(
                    GetFormulaEvaluationCacheKey(cell.CellReference!.Value!), out array!);
        }

        private sealed class FixedArraySheetIndex {
            internal readonly List<FixedArrayOwner> Owners = new List<FixedArrayOwner>();
            internal readonly List<FixedArrayOwner> DynamicOwners = new List<FixedArrayOwner>();
            internal readonly Dictionary<Cell, FixedArrayOwner> DynamicByCell = new Dictionary<Cell, FixedArrayOwner>();
            internal readonly List<Cell> FormulaCells = new List<Cell>();
            internal readonly Dictionary<long, Cell> Cells = new Dictionary<long, Cell>();
        }

        private sealed class FixedArrayOwner {
            internal Cell Cell = null!;
            internal int Top, Left, Bottom, Right;
            internal bool Dynamic;
        }

        private bool TryResolveFixedArrayChild(int row, int column, out FormulaArgumentValue result) {
            result = default;
            foreach (FixedArrayOwner owner in GetFixedArraySheetIndex().Owners) {
                if (row < owner.Top || row > owner.Bottom || column < owner.Left || column > owner.Right
                    || (row == owner.Top && column == owner.Left)) continue;
                result = TryEvaluateFormulaCellValue(owner.Cell, out _) && TryGetCalculatedArray(owner.Cell, out FormulaArrayValue evaluated)
                    ? evaluated.Values[(row - owner.Top) * evaluated.Columns + column - owner.Left]
                    : FormulaArgumentValue.UnresolvedFormula();
                return true;
            }
            return _formulaEvaluationCache != null && TryResolveDynamicArrayChild(row, column, out result);
        }

        private FixedArraySheetIndex GetFixedArraySheetIndex() {
            var sheets = ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache!).Sheets;
            if (!sheets.TryGetValue(Name, out FixedArraySheetIndex? index)) {
                index = new FixedArraySheetIndex();
                Metadata? metadata = null;
                foreach (Cell candidate in WorksheetRoot.Descendants<Cell>()) {
                    if (candidate.CellFormula != null) index.FormulaCells.Add(candidate);
                    if (candidate.CellFormula?.FormulaType?.Value == CellFormulaValues.Array
                        && IsSupportedArrayCall(candidate.CellFormula.Text)
                        && TryFixedArrayBounds(candidate.CellFormula.Reference?.Value ?? "",
                            out int top, out int left, out int bottom, out int right)) {
                        bool dynamic = false;
                        if (candidate.CellMetaIndex != null) {
                            var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
                            if (part != null && !part.IsRootElementLoaded)
                                ValidateInCellImageMetadataPart(part, "Cell metadata");
                            metadata ??= part?.Metadata;
                            dynamic = CreateFormulaArrayInfo(candidate, metadata)?.IsDynamic == true;
                        }
                        var owner = new FixedArrayOwner { Cell = candidate, Top = top, Left = left,
                            Bottom = bottom, Right = right, Dynamic = dynamic };
                        (dynamic ? index.DynamicOwners : index.Owners).Add(owner);
                        if (dynamic) index.DynamicByCell.Add(candidate, owner);
                    }
                }
                if (index.DynamicOwners.Count > 0)
                    foreach (Cell candidate in WorksheetRoot.Descendants<Cell>())
                        if (TryParseCellReference(candidate.CellReference?.Value ?? "", out int row, out int column))
                            index.Cells[DynamicCellKey(row, column)] = candidate;
                sheets[Name] = index;
            }
            return index;
        }
        private static long DynamicCellKey(int row, int column) => ((long)row << 15) | (uint)column;

        // Array-valued scalar expressions require a separate qualified contract.
        private sealed class FormulaArrayValue {
            internal FormulaArrayValue(int rows, int columns, FormulaArgumentValue[] values) {
                Rows = rows;
                Columns = columns;
                Values = values;
            }
            internal int Rows { get; }
            internal int Columns { get; }
            internal FormulaArgumentValue[] Values { get; }
        }

        private bool TryEvaluateAuthoredFormulaCell(Cell cell, string formula, out FormulaArgumentValue result) {
            result = default;
            if (cell.CellFormula?.FormulaType?.Value != CellFormulaValues.Array)
                return TryEvaluateFormulaValue(formula, out result);
            if (!IsSupportedArrayCall(formula)) return TryEvaluateFormulaValue(formula, out result);
            string? reference = cell.CellFormula.Reference?.Value;
            if (reference == null || !TryFixedArrayBounds(reference,
                out int r1, out int c1, out int r2, out int c2)
                || !TryParseCellReference(cell.CellReference?.Value ?? "", out int row, out int column)
                || row != r1 || column != c1
                || !TryGetFormulaRangeCellCount(r1, c1, r2, c2, out _)
                || !TryEvaluateArrayValue(formula, 0, out FormulaArrayValue array))
                return false;
            if (_formulaEvaluationCache != null
                && GetFixedArraySheetIndex().DynamicByCell.TryGetValue(cell, out FixedArrayOwner? dynamicOwner))
                return TryPlanDynamicArray(dynamicOwner, array, out result);
            if (array.Rows != r2 - r1 + 1 || array.Columns != c2 - c1 + 1) return false;
            // Imported or subsequently edited ranges can contain another formula.
            // Never replace that formula with an array cache.
            IEnumerable<Cell> formulaCells = _formulaEvaluationCache == null
                ? WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null)
                : GetFixedArraySheetIndex().FormulaCells;
            foreach (Cell occupied in formulaCells) {
                if (occupied == cell || occupied.CellFormula == null) continue;
                if (TryParseCellReference(occupied.CellReference?.Value ?? "", out int rr, out int cc)
                    && rr >= r1 && rr <= r2 && cc >= c1 && cc <= c2) return false;
            }
            foreach (MergeCell merge in WorksheetRoot.Elements<MergeCells>().SelectMany(m => m.Elements<MergeCell>())) {
                if (TryFixedArrayBounds(merge.Reference?.Value ?? "", out int mr1, out int mc1, out int mr2, out int mc2)
                    && RangesOverlapInclusive((r1, c1, r2, c2), (mr1, mc1, mr2, mc2))) return false;
            }
            result = array.Values[0];
            if (_formulaEvaluationCache != null)
                ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache).Results[
                    GetFormulaEvaluationCacheKey(cell.CellReference!.Value!)] = array;
            return true;
        }

        private static bool TryFixedArrayBounds(string reference, out int top, out int left, out int bottom, out int right) {
            reference = reference.Replace("$", "");
            if (A1.TryParseRange(reference, out top, out left, out bottom, out right)) return true;
            if (TryParseCellReference(reference, out top, out left)) { bottom = top; right = left; return true; }
            bottom = right = 0;
            return false;
        }

        private void WriteFixedArrayFormulaCache(Cell owner, FormulaArrayValue array) {
            var (row, column) = A1.ParseCellRef(owner.CellReference!.Value!);
            for (int r = 0; r < array.Rows; r++) {
                for (int c = 0; c < array.Columns; c++) {
                    Cell cell = r == 0 && c == 0 ? owner : GetCell(row + r, column + c);
                    cell.InlineString = null;
                    cell.ValueMetaIndex = null;
                    SetFormulaCachedValue(cell, array.Values[r * array.Columns + c]);
                }
            }
        }

        private static bool IsSupportedArrayCall(string expression) {
            if (!ExcelFormulaExpressionParser.TryParseFunctionCall(expression, out ExcelFormulaFunctionCallSyntax? call)) return false;
            string name = GetArrayFunctionName(call!.Name);
            return ExcelFormulaCapabilities.IsArrayFunction(name);
        }

        private static string GetArrayFunctionName(string name) {
            name = name.ToUpperInvariant();
            if (name.StartsWith("_XLFN.", StringComparison.Ordinal)) name = name.Substring(6);
            if (name.StartsWith("_XLWS.", StringComparison.Ordinal)) name = name.Substring(6);
            return name;
        }

        private bool TryEvaluateArrayValue(string expression, int depth, out FormulaArrayValue array) {
            array = null!;
            if (!HasSufficientFormulaExecutionStack() || depth >= 32 || expression.Length > MaxSupportedFormulaLength) return false;
            if (ExcelFormulaExpressionParser.TryParseComparison(expression, out ExcelFormulaBinaryExpressionSyntax? comparison)) {
                if (!TryEvaluateArrayValue(comparison!.Left, depth + 1, out FormulaArrayValue left)
                    || !TryEvaluateArrayValue(comparison.Right, depth + 1, out FormulaArrayValue right)) return false;
                bool leftScalar = left.Values.Length == 1, rightScalar = right.Values.Length == 1;
                if (!leftScalar && !rightScalar && (left.Rows != right.Rows || left.Columns != right.Columns)) return false;
                FormulaArrayValue shape = leftScalar ? right : left;
                var compared = new FormulaArgumentValue[shape.Values.Length];
                for (int i = 0; i < compared.Length; i++) {
                    var a = left.Values[leftScalar ? 0 : i];
                    var b = right.Values[rightScalar ? 0 : i];
                    if (a.IsError || b.IsError) compared[i] = a.IsError ? a : b;
                    else if (!a.HasValue || !b.HasValue || a.IsBoolean != b.IsBoolean
                        || a.Number.HasValue != b.Number.HasValue) return false;
                    else if (TryCompareFormulaValues(a, comparison.Operator, b, out bool matches))
                        compared[i] = new FormulaArgumentValue(matches ? 1 : 0, null, isBoolean: true);
                    else return false;
                }
                array = new FormulaArrayValue(shape.Rows, shape.Columns, compared);
                return true;
            }
            if (TryResolveFormulaRangeReference(expression, out ExcelSheet sheet,
                out int r1, out int c1, out int r2, out int c2)) {
                if (!TryGetFormulaRangeCellCount(r1, c1, r2, c2, out int count)) return false;
                var values = new FormulaArgumentValue[count];
                int index = 0;
                for (int r = r1; r <= r2; r++) {
                    for (int c = c1; c <= c2; c++) {
                        FormulaArgumentValue value = sheet.ResolveCellArgument(r, c);
                        if (value.IsUnresolvedFormula) return false;
                        values[index++] = value;
                    }
                }
                array = new FormulaArrayValue(r2 - r1 + 1, c2 - c1 + 1, values);
                return true;
            }
            if (!ExcelFormulaExpressionParser.TryParseFunctionCall(expression, out ExcelFormulaFunctionCallSyntax? call)) {
                if (!TryResolveFormulaArgument(expression, out FormulaArgumentValue scalar) || scalar.IsUnresolvedFormula) return false;
                array = new FormulaArrayValue(1, 1, new[] { scalar.HasValue ? scalar : new FormulaArgumentValue(0, null) });
                return true;
            }
            var args = SplitFormulaArguments(call!.Arguments);
            string function = GetArrayFunctionName(call.Name);
            if (function == "SEQUENCE") return TryEvaluateSequence(args, out array);
            if (args.Count < 1 || !TryEvaluateArrayValue(args[0], depth + 1, out FormulaArrayValue input)) return false;
            if (function == "FILTER") return TryEvaluateArrayFilter(args, input, depth, out array);
            if (function == "SORT") return TryEvaluateArraySort(args, input, out array);
            if (function == "UNIQUE") return TryEvaluateArrayUnique(args, input, out array);
            return false;
        }

        private bool TryArrayNumber(IReadOnlyList<string> args, int index, double fallback, out double number) {
            number = fallback;
            if (index >= args.Count || string.IsNullOrWhiteSpace(args[index])) return true;
            if (!TryResolveFormulaArgument(args[index], out FormulaArgumentValue value)
                || value.IsUnresolvedFormula || value.IsError) return false;
            if (value.Number.HasValue) number = value.Number.Value;
            else if (!double.TryParse(value.Text, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out number)) return false;
            return !double.IsNaN(number) && !double.IsInfinity(number);
        }

        private bool TryEvaluateSequence(IReadOnlyList<string> args, out FormulaArrayValue array) {
            array = null!;
            if (args.Count < 1 || args.Count > 4
                || !TryArrayNumber(args, 0, 1, out double rows)
                || !TryArrayNumber(args, 1, 1, out double columns)
                || !TryArrayNumber(args, 2, 1, out double start)
                || !TryArrayNumber(args, 3, 1, out double step)) return false;
            rows = Math.Truncate(rows);
            columns = Math.Truncate(columns);
            if (rows < 1 || columns < 1) {
                array = ArrayError(rows < 0 || columns < 0 ? "#VALUE!" : "#CALC!");
                return true;
            }
            if (rows > A1.MaxRows || columns > A1.MaxColumns || rows * columns > MaxResolvedFormulaRangeCells) return false;
            var values = new FormulaArgumentValue[(int)(rows * columns)];
            for (int i = 0; i < values.Length; i++) {
                double value = start + i * step;
                if (double.IsInfinity(value) || double.IsNaN(value)) {
                    array = ArrayError("#NUM!");
                    return true;
                }
                values[i] = new FormulaArgumentValue(value, null);
            }
            array = new FormulaArrayValue((int)rows, (int)columns, values);
            return true;
        }

        private static FormulaArrayValue ArrayError(string error) =>
            new FormulaArrayValue(1, 1, new[] { FormulaArgumentValue.Error(error) });

        private bool TryEvaluateArrayFilter(IReadOnlyList<string> args, FormulaArrayValue input,
            int depth, out FormulaArrayValue array) {
            array = null!;
            if (args.Count < 2 || args.Count > 3
                || !TryEvaluateArrayValue(args[1], depth + 1, out FormulaArrayValue include)) return false;
            bool byRows = include.Columns == 1 && include.Rows == input.Rows;
            bool byColumns = include.Rows == 1 && include.Columns == input.Columns;
            if (!byRows && !byColumns) { array = ArrayError("#VALUE!"); return true; }
            var selected = new List<int>();
            for (int i = 0; i < include.Values.Length; i++) {
                var value = include.Values[i];
                if (value.IsError) { array = ArrayError(value.ErrorCode!); return true; }
                if (!value.HasValue) continue;
                if (!value.Number.HasValue) return false;
                if (value.Number.Value != 0) selected.Add(i);
            }
            if (selected.Count == 0) {
                if (args.Count == 2) { array = ArrayError("#CALC!"); return true; }
                if (!TryResolveFormulaArgument(args[2], out FormulaArgumentValue empty) || empty.IsUnresolvedFormula) return false;
                array = new FormulaArrayValue(1, 1, new[] { empty.HasValue ? empty : new FormulaArgumentValue(0, null) });
                return true;
            }
            array = SelectArrayVectors(input, selected, !byRows);
            return true;
        }

        private bool TryEvaluateArraySort(IReadOnlyList<string> args, FormulaArrayValue input, out FormulaArrayValue array) {
            array = null!;
            if (args.Count > 4 || !TryArrayNumber(args, 1, 1, out double index)
                || !TryArrayNumber(args, 2, 1, out double order)
                || !TryArrayFlag(args, 3, out bool byColumns)) return false;
            int keyCount = byColumns ? input.Rows : input.Columns;
            index = Math.Truncate(index);
            if (index < 1 || index > keyCount || (order != 1 && order != -1)) {
                array = ArrayError("#VALUE!"); return true;
            }
            int vectors = byColumns ? input.Columns : input.Rows;
            int key = (int)index - 1;
            var keyKinds = new ArraySortKeyKind[vectors];
            for (int i = 0; i < vectors; i++) {
                var value = input.Values[byColumns ? key * input.Columns + i : i * input.Columns + key];
                if (!TryGetArraySortKeyKind(value, out keyKinds[i])) return false;
            }
            var selected = Enumerable.Range(0, vectors).ToArray();
            System.Array.Sort(selected, (left, right) => {
                ArraySortKeyKind leftKind = keyKinds[left], rightKind = keyKinds[right];
                int comparison;
                if (leftKind == ArraySortKeyKind.Blank || rightKind == ArraySortKeyKind.Blank) {
                    // Excel keeps blank keys last in both sort directions.
                    comparison = leftKind == rightKind ? 0 : leftKind == ArraySortKeyKind.Blank ? 1 : -1;
                } else {
                    comparison = ((int)leftKind).CompareTo((int)rightKind);
                    if (comparison == 0) {
                        FormulaArgumentValue leftValue = input.Values[byColumns ? key * input.Columns + left : left * input.Columns + key];
                        FormulaArgumentValue rightValue = input.Values[byColumns ? key * input.Columns + right : right * input.Columns + key];
                        comparison = leftKind == ArraySortKeyKind.Text
                            ? string.Compare(leftValue.Text, rightValue.Text, StringComparison.OrdinalIgnoreCase)
                            : leftValue.Number!.Value.CompareTo(rightValue.Number!.Value);
                    }
                    comparison = Math.Sign(comparison) * (int)order;
                }
                return comparison == 0 ? left.CompareTo(right) : comparison;
            });
            array = SelectArrayVectors(input, selected, byColumns);
            return true;
        }

        private enum ArraySortKeyKind : byte { Number, Text, Boolean, Blank }

        private static bool TryGetArraySortKeyKind(FormulaArgumentValue value, out ArraySortKeyKind kind) {
            kind = default;
            if (value.IsError || value.IsUnresolvedFormula) return false;
            if (!value.HasValue) { kind = ArraySortKeyKind.Blank; return true; }
            if (value.IsBoolean) { kind = ArraySortKeyKind.Boolean; return value.Number.HasValue; }
            if (value.Number.HasValue) {
                if (double.IsNaN(value.Number.Value) || double.IsInfinity(value.Number.Value)) return false;
                kind = ArraySortKeyKind.Number;
                return true;
            }
            if (value.Text == null || value.Text.Length == 0) return false;
            foreach (char character in value.Text) {
                if (!((character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z')))
                    return false;
            }
            kind = ArraySortKeyKind.Text;
            return true;
        }

        private bool TryArrayFlag(IReadOnlyList<string> args, int index, out bool flag) {
            flag = false;
            if (!TryArrayNumber(args, index, 0, out double number)) return false;
            flag = number != 0;
            return true;
        }

        private bool TryEvaluateArrayUnique(IReadOnlyList<string> args, FormulaArrayValue input, out FormulaArrayValue array) {
            array = null!;
            if (args.Count > 3 || !TryArrayFlag(args, 1, out bool byColumns)
                || !TryArrayFlag(args, 2, out bool exactlyOnce)) return false;
            int vectors = byColumns ? input.Columns : input.Rows;
            var comparer = new ArrayVectorComparer(input, byColumns);
            var counts = new Dictionary<int, int>(comparer);
            for (int i = 0; i < vectors; i++) {
                counts.TryGetValue(i, out int count);
                counts[i] = count + 1;
            }
            var selected = new List<int>();
            var seen = new HashSet<int>(comparer);
            for (int i = 0; i < vectors; i++)
                if (seen.Add(i) && (!exactlyOnce || counts[i] == 1)) selected.Add(i);
            if (selected.Count == 0) { array = ArrayError("#CALC!"); return true; }
            array = SelectArrayVectors(input, selected, byColumns);
            return true;
        }

        private static FormulaArrayValue SelectArrayVectors(FormulaArrayValue input, IReadOnlyList<int> selected, bool byColumns) {
            int rows = byColumns ? input.Rows : selected.Count;
            int columns = byColumns ? selected.Count : input.Columns;
            var values = new FormulaArgumentValue[rows * columns];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < columns; c++)
                    values[r * columns + c] = NormalizeArrayOutput(input.Values[(byColumns ? r : selected[r]) * input.Columns + (byColumns ? selected[c] : c)]);
            return new FormulaArrayValue(rows, columns, values);
        }

        private sealed class ArrayVectorComparer : IEqualityComparer<int> {
            private readonly FormulaArrayValue _input;
            private readonly bool _byColumns;
            internal ArrayVectorComparer(FormulaArrayValue input, bool byColumns) { _input = input; _byColumns = byColumns; }
            private FormulaArgumentValue Value(int vector, int offset) =>
                NormalizeArrayOutput(_input.Values[_byColumns ? offset * _input.Columns + vector : vector * _input.Columns + offset]);
            private int Length => _byColumns ? _input.Rows : _input.Columns;
            public bool Equals(int left, int right) {
                for (int i = 0; i < Length; i++) {
                    var a = Value(left, i);
                    var b = Value(right, i);
                    if (a.IsBoolean != b.IsBoolean || a.IsError != b.IsError || a.Number != b.Number
                        || (!a.Number.HasValue && !string.Equals(a.Text, b.Text, StringComparison.OrdinalIgnoreCase))) return false;
                }
                return true;
            }
            public int GetHashCode(int vector) {
                unchecked {
                    int hash = 17;
                    for (int i = 0; i < Length; i++) {
                        var value = Value(vector, i);
                        hash = hash * 31 + (value.Number?.GetHashCode() ?? StringComparer.OrdinalIgnoreCase.GetHashCode(value.Text ?? ""));
                        hash = hash * 31 + (value.IsBoolean ? 1 : value.IsError ? 2 : 0);
                    }
                    return hash;
                }
            }
        }

        private static FormulaArgumentValue NormalizeArrayOutput(FormulaArgumentValue value) =>
            value.HasValue ? value : new FormulaArgumentValue(0, null);
    }
}
