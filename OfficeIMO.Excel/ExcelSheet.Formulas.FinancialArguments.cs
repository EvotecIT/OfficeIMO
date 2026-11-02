using System.Globalization;

namespace OfficeIMO.Excel {

    public partial class ExcelSheet {
        private bool TryFinancialScalar(string token, out double number, out FormulaArgumentValue error) {
            number = 0; error = default;
            if (!TryResolveFormulaArgument(token, out var value) || value.IsUnresolvedFormula) return false;
            if (value.IsError) { error = value; return true; }
            if (!value.HasValue) return true;
            if (value.Number.HasValue) number = value.Number.Value;
            else if (!double.TryParse(value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) {
                error = FormulaArgumentValue.Error("#VALUE!"); return true;
            }
            if (!IsFinite(number)) error = FormulaArgumentValue.Error("#NUM!");
            return true;
        }

        private bool TryResolveFinancialSeries(string token, ref int budget, out List<FormulaArgumentValue> values) {
            if (TryResolveFormulaRange(token, out values, ref budget, out bool isReference)) return true;
            if (isReference) return false;
            token = token.Trim();
            values = new List<FormulaArgumentValue>();
            if (token.Length < 2 || token[0] != '{' || token[token.Length - 1] != '}') {
                if (!TryResolveFormulaArgument(token, out var scalar) || budget-- <= 0) return false;
                values.Add(scalar); return true;
            }
            return TryResolveFormulaArrayConstant(token, ref budget, out values, out _);
        }
    }
}
