namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string NormalizeSupportedFunctionPrefix(string formula) {
            const string futurePrefix = "_xlfn.";
            if (formula.IndexOf(futurePrefix, StringComparison.OrdinalIgnoreCase) < 0
                || !ExcelFormulaExpressionParser.TryParseFunctionCall(formula, out ExcelFormulaFunctionCallSyntax? call)) {
                return formula;
            }

            string storedName = call!.Name;
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
