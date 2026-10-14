namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
