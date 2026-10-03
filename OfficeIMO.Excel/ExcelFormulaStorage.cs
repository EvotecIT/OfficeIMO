using System.Text;

namespace OfficeIMO.Excel;

/// <summary>Prepares qualified formula functions for their XLSX storage spelling.</summary>
internal static class ExcelFormulaStorage {
    /// <summary>Prefixes qualified TEXTJOIN calls while preserving literals, references, and authored prefixes.</summary>
    internal static string PrepareForXlsx(string formula) {
        if (formula.IndexOf("TEXTJOIN", StringComparison.OrdinalIgnoreCase) < 0) return formula;
        ExcelFormulaSyntaxTree syntax = ExcelFormulaSyntaxTree.Parse(formula);
        var result = new StringBuilder(formula.Length + 6);
        bool changed = false;
        foreach (ExcelFormulaSyntaxNode node in syntax.Nodes) {
            if (node is ExcelFormulaFunctionSyntax function
                && function.Name.Equals("TEXTJOIN", StringComparison.OrdinalIgnoreCase)) {
                result.Append("_xlfn.");
                changed = true;
            }
            result.Append(node.Text);
        }
        return changed ? result.ToString() : formula;
    }
}
