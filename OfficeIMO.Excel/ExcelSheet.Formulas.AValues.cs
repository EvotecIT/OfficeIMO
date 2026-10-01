namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    private bool TryEvaluateAValueAggregate(string function, string args, out FormulaArgumentValue result) {
        result = default;
        int argumentCount = SplitFormulaArguments(args).Count;
        if (argumentCount < 1 || argumentCount > 255
            || !TryResolveFormulaArguments(args, out var values)
            || values.Any(value => value.IsUnresolvedFormula)) return false;
        foreach (FormulaArgumentValue value in values) {
            if (value.IsError) { result = value; return true; }
        }

        var numbers = new List<double>();
        foreach (FormulaArgumentValue value in values) {
            if (value.SourceCellKind == ExcelCellDataKind.Boolean) {
                numbers.Add(value.Number.HasValue ? (value.Number.Value != 0 ? 1d : 0d)
                    : value.Text == "1" || (bool.TryParse(value.Text, out bool logical) && logical) ? 1d : 0d);
            } else if (value.SourceCellKind == ExcelCellDataKind.Text) {
                numbers.Add(0d);
            } else if (value.Number.HasValue) {
                numbers.Add(value.Number.Value);
            } else if (value.Text != null) {
                numbers.Add(0d);
            }
        }
        if (numbers.Count == 0) {
            result = function == "AVERAGEA" ? FormulaArgumentValue.Error("#DIV/0!") : new FormulaArgumentValue(0d, "0");
            return true;
        }
        double number = function == "AVERAGEA" ? numbers.Average() : function == "MINA" ? numbers.Min() : numbers.Max();
        if (!IsFinite(number)) return false;
        result = new FormulaArgumentValue(number, InvariantNumberText.Get(number));
        return true;
    }
}
