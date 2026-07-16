using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string NormalizeSupportedFunctionPrefix(string formula) {
            Match match = FunctionNameFormulaRegex.Match(formula);
            if (!match.Success) {
                return formula;
            }

            string storedName = match.Groups[1].Value;
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

            Group nameGroup = match.Groups[1];
            return formula.Remove(nameGroup.Index, nameGroup.Length).Insert(nameGroup.Index, functionName);
        }
    }
}
