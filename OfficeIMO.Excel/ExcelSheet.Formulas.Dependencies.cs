using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private string GetUnsupportedFormulaReason(string formula) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return "Formula is empty.";
            }

            if (formula.Length > MaxSupportedFormulaLength) {
                return $"Formula is longer than {MaxSupportedFormulaLength} characters.";
            }

            try {
                if (formula.IndexOf(';') >= 0) {
                    return "Formula uses semicolon argument separators; OfficeIMO's lightweight evaluator expects Open XML comma-separated formulas.";
                }

                if (formula.IndexOf('&') >= 0) {
                    return "Formula uses the text concatenation operator, which OfficeIMO's lightweight evaluator does not currently support.";
                }

                if (formula.IndexOf('{') >= 0 || formula.IndexOf('}') >= 0) {
                    return "Formula uses array constants, which OfficeIMO's lightweight evaluator does not currently support.";
                }

                Match supportedFunctionMatch = SimpleFunctionFormulaRegex.Match(formula);
                if (supportedFunctionMatch.Success) {
                    string function = supportedFunctionMatch.Groups[1].Value.ToUpperInvariant();
                    return $"Formula uses supported function '{function}' with arguments OfficeIMO's lightweight evaluator cannot currently evaluate.";
                }

                Match functionMatch = FunctionNameFormulaRegex.Match(formula);
                if (functionMatch.Success) {
                    string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                    if (_excelDocument.Calculation.TryGetCustomFunction(function, out _)) {
                        return $"Formula uses registered custom function '{function}' with arguments the custom evaluator cannot currently evaluate.";
                    }

                    return $"Function '{function}' is not supported by OfficeIMO's lightweight evaluator.";
                }
            } catch (RegexMatchTimeoutException) {
                return "Formula diagnostics timed out while parsing the formula.";
            }

            return "Formula is outside OfficeIMO's lightweight evaluator support.";
        }

        private IReadOnlyList<string> GetFormulaDependencies(string formula) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return Array.Empty<string>();
            }

            try {
                string searchableFormula = MaskFormulaStringLiterals(formula);
                return FormulaReferenceRegex.Matches(searchableFormula)
                    .Cast<Match>()
                    .Select(match => match.Groups["reference"].Value)
                    .Where(reference => !string.IsNullOrWhiteSpace(reference))
                    .Select(NormalizeFormulaDependencyReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(reference => reference, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            } catch (RegexMatchTimeoutException) {
                return Array.Empty<string>();
            }
        }

        private IReadOnlyList<string> GetFormulaDependencyIssues(string formula, string? sourceCellReference, IReadOnlyList<string> dependencies) {
            if (dependencies.Count == 0) {
                return Array.Empty<string>();
            }

            var issues = new List<string>();
            string? sourceReference = NormalizeFormulaCellReference(sourceCellReference);
            foreach (string dependency in dependencies) {
                if (!TryResolveFormulaRangeReference(dependency, out ExcelSheet dependencySheet, out int r1, out int c1, out int r2, out int c2)) {
                    issues.Add($"Cannot resolve dependency '{dependency}'.");
                    continue;
                }

                if (sourceReference != null
                    && string.Equals(dependencySheet.Name, Name, StringComparison.OrdinalIgnoreCase)
                    && TryParseCellReference(sourceReference, out int sourceRow, out int sourceColumn)
                    && sourceRow >= r1 && sourceRow <= r2 && sourceColumn >= c1 && sourceColumn <= c2) {
                    issues.Add($"Dependency '{dependency}' references its own formula cell.");
                }

                foreach (Cell dependencyCell in dependencySheet.WorksheetRoot.Descendants<Cell>().Where(cell => cell.CellFormula != null)) {
                    string? dependencyReference = NormalizeFormulaCellReference(dependencyCell.CellReference?.Value);
                    if (dependencyReference == null
                        || !TryParseCellReference(dependencyReference, out int dependencyRow, out int dependencyColumn)
                        || dependencyRow < r1 || dependencyRow > r2 || dependencyColumn < c1 || dependencyColumn > c2) {
                        continue;
                    }

                    string formattedDependencyCell = $"{dependencySheet.Name}!{dependencyReference}";
                    if (string.Equals(dependencySheet.Name, Name, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(dependencyReference, sourceReference, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    if (dependencyCell.CellValue == null) {
                        issues.Add($"Dependency '{formattedDependencyCell}' is a formula without a cached result.");
                    }

                    string dependencyFormula = dependencyCell.CellFormula!.Text ?? string.Empty;
                    if (!TryEvaluateFormulaValue(dependencyFormula, out _)) {
                        issues.Add($"Dependency '{formattedDependencyCell}' contains a formula outside the lightweight evaluator support.");
                    }
                }
            }

            return issues
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(issue => issue, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string NormalizeFormulaDependencyReference(string reference) {
            string normalized = reference.Trim().Replace("$", string.Empty);
            if (TryResolveFormulaRangeReference(normalized, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2)) {
                string start = A1.CellReference(r1, c1);
                string end = A1.CellReference(r2, c2);
                return r1 == r2 && c1 == c2
                    ? $"{sheet.Name}!{start}"
                    : $"{sheet.Name}!{start}:{end}";
            }

            return normalized;
        }

        private static string MaskFormulaStringLiterals(string formula) {
            var builder = new StringBuilder(formula.Length);
            bool inString = false;
            for (int i = 0; i < formula.Length; i++) {
                char character = formula[i];
                if (character == '"') {
                    builder.Append(' ');
                    if (inString && i + 1 < formula.Length && formula[i + 1] == '"') {
                        i++;
                        builder.Append(' ');
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                builder.Append(inString ? ' ' : character);
            }

            return builder.ToString();
        }
    }
}
