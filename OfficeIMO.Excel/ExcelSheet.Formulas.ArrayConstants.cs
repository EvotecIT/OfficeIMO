using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // One rectangular, literal-only array parser shared by calculation owners.
        private static bool TryResolveFormulaArrayConstant(string token, ref int budget, out FormulaArrayValue array) {
            array = null!;
            if (!TryResolveFormulaArrayConstant(token, ref budget, out List<FormulaArgumentValue> values, out int columns)) return false;
            array = new FormulaArrayValue(values.Count / columns, columns, values.ToArray());
            return true;
        }

        private static bool TryResolveFormulaArrayConstant(string token, ref int budget,
            out List<FormulaArgumentValue> values, out int columnCount) {
            values = new List<FormulaArgumentValue>();
            columnCount = 0;
            token = token.Trim();
            if (token.Length < 2 || token[0] != '{' || token[token.Length - 1] != '}') return false;
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
                if (double.TryParse(literal, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    if (!IsFinite(number)) return false;
                    value = new(number, literal);
                }
                else if (ExcelFormulaExpressionParser.TryParseTextLiteral(literal, out string text)) value = new(null, text);
                else if (literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || literal.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                    value = FormulaArgumentValue.Boolean(literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase));
                else if (TryParseFormulaErrorLiteral(literal, out string error)) value = FormulaArgumentValue.Error(error);
                else return false;
                if (budget-- <= 0) return false;
                values.Add(value);
                columns++;
                start = index + 1;
                if (ch == ';' || index == token.Length - 1) {
                    if (width >= 0 && columns != width) return false;
                    width = columns;
                    columns = 0;
                }
            }
            if (quoted || width <= 0 || values.Count == 0) return false;
            columnCount = width;
            return true;
        }
    }
}
