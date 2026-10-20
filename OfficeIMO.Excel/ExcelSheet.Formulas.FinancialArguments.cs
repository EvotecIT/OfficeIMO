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
            // Array constants contain literals only; do not turn them into a second expression evaluator.
            int start = 1, width = -1, columns = 0;
            bool quoted = false;
            for (int index = 1; index < token.Length; index++) {
                char ch = token[index];
                if (ch == '"') {
                    if (quoted && index + 1 < token.Length && token[index + 1] == '"') index++;
                    else quoted = !quoted;
                    continue;
                }
                if (quoted || ch != ',' && ch != ';' && index != token.Length - 1) continue;
                string literal = token.Substring(start, index - start).Trim();
                FormulaArgumentValue value;
                if (double.TryParse(literal, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) value = new(number, literal);
                else if (ExcelFormulaExpressionParser.TryParseTextLiteral(literal, out string text)) value = new(null, text);
                else if (literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || literal.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                    value = FormulaArgumentValue.Boolean(literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase));
                else if (TryParseFormulaErrorLiteral(literal, out string error)) value = FormulaArgumentValue.Error(error);
                else return false;
                if (budget-- <= 0) return false;
                values.Add(value); columns++; start = index + 1;
                if (ch == ';' || index == token.Length - 1) {
                    if (width >= 0 && columns != width) return false;
                    width = columns; columns = 0;
                }
            }
            return !quoted && values.Count > 0;
        }
    }
}
