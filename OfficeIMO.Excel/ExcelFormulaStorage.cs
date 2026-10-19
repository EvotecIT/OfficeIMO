using System.Text;

namespace OfficeIMO.Excel;

/// <summary>Prepares qualified formula functions for their XLSX storage spelling.</summary>
internal static class ExcelFormulaStorage {
    private static readonly string[] QualifiedFutureFunctions = {
        "TEXTJOIN", "IFNA", "MINIFS", "MAXIFS",
        "IFS", "CONCAT", "TEXTBEFORE", "TEXTAFTER", "REGEXEXTRACT",
        "XLOOKUP", "XMATCH", "SWITCH", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P",
        "CEILING.MATH", "FLOOR.MATH", "COVARIANCE.P", "COVARIANCE.S",
        "RANK.EQ", "RANK.AVG", "MODE.SNGL", "PERCENTILE.INC", "PERCENTILE.EXC",
        "PERCENTRANK.INC", "PERCENTRANK.EXC", "QUARTILE.INC", "QUARTILE.EXC",
        "FORECAST.LINEAR", "FORMULATEXT", "ISFORMULA", "ISOWEEKNUM", "DAYS"
    };
    private static readonly string[] QualifiedArrayFunctions = { "SEQUENCE", "UNIQUE", "SORT", "FILTER" };

    /// <summary>Prefixes qualified future-function calls while preserving literals, references, and authored prefixes.</summary>
    internal static string PrepareForXlsx(string formula, bool qualifyArrayFunctions = false) {
        if (!MayContainFunction(formula, QualifiedFutureFunctions)
            && (!qualifyArrayFunctions || !MayContainFunction(formula, QualifiedArrayFunctions))) return formula;
        ExcelFormulaSyntaxTree syntax = ExcelFormulaSyntaxTree.Parse(formula);
        var result = new StringBuilder(formula.Length + 6);
        bool changed = false;
        int offset = 0;
        foreach (ExcelFormulaSyntaxNode node in syntax.Nodes) {
            if (node is ExcelFormulaFunctionSyntax function
                && (offset == 0 || formula[offset - 1] != '!')
                && StoragePrefix(function.Name, qualifyArrayFunctions) is string prefix) {
                result.Append(prefix);
                changed = true;
            }
            result.Append(node.Text);
            offset += node.Text.Length;
        }
        return changed ? result.ToString() : formula;
    }

    private static bool MayContainFunction(string formula, string[] functions) {
        foreach (string name in functions)
            if (formula.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        return false;
    }

    private static string? StoragePrefix(string function, bool qualifyArrayFunctions) {
        foreach (string name in QualifiedFutureFunctions)
            if (function.Equals(name, StringComparison.OrdinalIgnoreCase)) return "_xlfn.";
        if (qualifyArrayFunctions) {
            foreach (string name in QualifiedArrayFunctions) {
                if (function.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return name == "SORT" || name == "FILTER" ? "_xlfn._xlws." : "_xlfn.";
            }
        }
        return null;
    }
}
