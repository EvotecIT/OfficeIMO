namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluatePivotDataValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || tokens.Count > 514 || (tokens.Count & 1) != 0
                || !TryResolveFormulaArgument(tokens[0], out var dataField)) return false;
            if (dataField.IsError) { result = dataField; return true; }
            if (dataField.Number.HasValue || dataField.Text == null) return false;
            var criteria = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (int index = 2; index < tokens.Count; index += 2) {
                if (!TryResolveFormulaArgument(tokens[index], out var field) || !TryResolveFormulaArgument(tokens[index + 1], out var item)) return false;
                if (field.IsError || item.IsError) { result = field.IsError ? field : item; return true; }
                if (field.Number.HasValue || field.Text == null) return false;
                object? key = item.IsBoolean ? item.Number != 0 : item.Number.HasValue ? (object)item.Number.Value : item.Text;
                if (criteria.TryGetValue(field.Text, out object? previous) && !PivotLookupValuesEqual(previous, key)) {
                    result = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                criteria[field.Text] = key;
            }
            if (!TryResolveFormulaRangeReference(tokens[1], out var sheet, out int top, out int left, out int bottom, out int right)) {
                result = FormulaArgumentValue.Error("#REF!");
                return true;
            }
            switch (sheet.ReadSavedPivotDataAt(top, left, bottom, right, dataField.Text, criteria, out ExcelCellData value)) {
                case ExcelPivotDataLookup.NotFound: result = FormulaArgumentValue.Error("#REF!"); return true;
                case ExcelPivotDataLookup.Unsupported: return false;
            }
            result = value.Kind == ExcelCellDataKind.Error ? FormulaArgumentValue.Error(value.Value as string ?? "#VALUE!")
                : value.Value is bool boolean ? new FormulaArgumentValue(boolean ? 1 : 0, boolean ? "TRUE" : "FALSE", isBoolean: true)
                : value.Value is double number ? new FormulaArgumentValue(number, InvariantNumberText.Get(number))
                : new FormulaArgumentValue(null, value.Value as string);
            return result.HasValue;
        }
    }
}
