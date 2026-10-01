namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Native Excel requires these future-function aliases in authored XLSX formulas.
        // Preserve literals, qualified names, structured references and existing aliases.
        private static string QualifyAuthoredArrayFunctions(string formula) {
            System.Text.StringBuilder? output = null;
            int copied = 0, brackets = 0;
            char quote = '\0';
            for (int index = 0; index < formula.Length; index++) {
                char current = formula[index];
                if (quote != '\0') {
                    if (current == quote) {
                        if (index + 1 < formula.Length && formula[index + 1] == quote) index++;
                        else quote = '\0';
                    }
                    continue;
                }
                if (brackets > 0) {
                    if (current == '\'' && index + 1 < formula.Length) { index++; continue; }
                    if (current == '[') brackets++;
                    else if (current == ']') brackets--;
                    continue;
                }
                if (current == '"' || current == '\'') { quote = current; continue; }
                if (current == '[') { brackets++; continue; }
                if (!char.IsLetter(current) && current != '_') continue;
                int start = index;
                while (index + 1 < formula.Length && (char.IsLetterOrDigit(formula[index + 1])
                    || formula[index + 1] == '_' || formula[index + 1] == '.')) index++;
                int end = index + 1, opening = end;
                while (opening < formula.Length && char.IsWhiteSpace(formula[opening])) opening++;
                if (opening >= formula.Length || formula[opening] != '('
                    || (start > 0 && (formula[start - 1] == '!' || formula[start - 1] == '\\'))) continue;
                string name = formula.Substring(start, end - start);
                string? prefix = name.Equals("SEQUENCE", StringComparison.OrdinalIgnoreCase)
                    || name.Equals("UNIQUE", StringComparison.OrdinalIgnoreCase) ? "_xlfn."
                    : name.Equals("SORT", StringComparison.OrdinalIgnoreCase)
                    || name.Equals("FILTER", StringComparison.OrdinalIgnoreCase) ? "_xlfn._xlws." : null;
                if (prefix == null) continue;
                output ??= new System.Text.StringBuilder(formula.Length + 16);
                output.Append(formula, copied, start - copied).Append(prefix);
                copied = start;
            }
            return output == null ? formula : output.Append(formula, copied, formula.Length - copied).ToString();
        }

        private static string NormalizeSupportedFunctionPrefix(string formula) {
            if (!ExcelFormulaExpressionParser.TryParseFunctionCall(formula, out ExcelFormulaFunctionCallSyntax? call)) {
                return formula;
            }

            string storedName = call!.Name;
            const string futurePrefix = "_xlfn.";
            if (!storedName.StartsWith(futurePrefix, StringComparison.OrdinalIgnoreCase)) {
                return formula;
            }

            string functionName = storedName.Substring(futurePrefix.Length);
            const string worksheetPrefix = "_xlws.";
            if (functionName.StartsWith(worksheetPrefix, StringComparison.OrdinalIgnoreCase)) {
                functionName = functionName.Substring(worksheetPrefix.Length);
            }

            if (!ExcelFormulaCapabilities.IsBuiltInFunction(functionName)) {
                return formula;
            }

            return formula.Remove(call.NameStart, call.NameLength).Insert(call.NameStart, functionName);
        }
    }
}
