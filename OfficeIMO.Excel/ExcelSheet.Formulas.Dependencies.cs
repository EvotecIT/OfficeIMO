using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaDependencyAlias {
            internal FormulaDependencyAlias(string text, bool allowStructuredSuffix) {
                Text = text;
                AllowStructuredSuffix = allowStructuredSuffix;
            }

            internal string Text { get; }
            internal bool AllowStructuredSuffix { get; }
        }

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

        private IReadOnlyList<string> GetFormulaDependencies(
            string formula,
            IReadOnlyList<FormulaDependencyAlias> aliases) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return Array.Empty<string>();
            }

            try {
                string searchableFormula = MaskFormulaStringLiterals(formula);
                var dependencies = new HashSet<string>(FormulaReferenceRegex.Matches(searchableFormula)
                    .Cast<Match>()
                    .Select(match => match.Groups["reference"].Value)
                    .Where(reference => !string.IsNullOrWhiteSpace(reference))
                    .Select(NormalizeFormulaDependencyReference), StringComparer.OrdinalIgnoreCase);
                foreach (FormulaDependencyAlias alias in aliases) {
                    AddFormulaAliasDependencies(
                        searchableFormula,
                        alias.Text,
                        alias.AllowStructuredSuffix,
                        dependencies);
                }

                return dependencies
                    .OrderBy(reference => reference, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            } catch (RegexMatchTimeoutException) {
                return Array.Empty<string>();
            }
        }

        private IReadOnlyList<FormulaDependencyAlias> GetFormulaDependencyAliases() {
            var aliases = new List<FormulaDependencyAlias>();
            var aliasKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            DefinedNames? definedNames = WorkbookRoot.DefinedNames;
            List<Sheet> sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int currentSheetIndex = sheets.FindIndex(sheet => string.Equals(sheet.Name?.Value, Name, StringComparison.OrdinalIgnoreCase));
            if (definedNames != null) {
                foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                    string? name = definedName.Name?.Value;
                    if (!IsFormulaDefinedNameToken(name ?? string.Empty) || IsBuiltInFormulaDefinedName(name)) {
                        continue;
                    }

                    if (definedName.LocalSheetId?.Value is uint localSheetIndex) {
                        if (localSheetIndex >= (uint)sheets.Count) {
                            continue;
                        }

                        if (localSheetIndex == (uint)currentSheetIndex) {
                            AddFormulaDependencyAlias(aliases, aliasKeys, name!, allowStructuredSuffix: false);
                        }

                        if (sheets[(int)localSheetIndex].Name?.Value is string scopedSheetName) {
                            AddFormulaDependencyAlias(aliases, aliasKeys, scopedSheetName + "!" + name, allowStructuredSuffix: false);
                            AddFormulaDependencyAlias(
                                aliases,
                                aliasKeys,
                                "'" + scopedSheetName.Replace("'", "''") + "'!" + name,
                                allowStructuredSuffix: false);
                        }

                        continue;
                    }

                    AddFormulaDependencyAlias(aliases, aliasKeys, name!, allowStructuredSuffix: false);
                }
            }

            WorkbookPart? workbookPart = _spreadSheetDocument.WorkbookPart;
            if (workbookPart != null) {
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                    foreach (Table table in worksheetPart.TableDefinitionParts.Select(part => part.Table).OfType<Table>()) {
                        IEnumerable<string> tableAliases = new[] { table.Name?.Value, table.DisplayName?.Value }
                            .OfType<string>()
                            .Where(alias => !string.IsNullOrWhiteSpace(alias))
                            .Distinct(StringComparer.OrdinalIgnoreCase);
                        foreach (string alias in tableAliases) {
                            AddFormulaDependencyAlias(aliases, aliasKeys, alias, allowStructuredSuffix: true);
                        }
                    }
                }
            }

            return aliases;
        }

        private static void AddFormulaDependencyAlias(
            ICollection<FormulaDependencyAlias> aliases,
            ISet<string> aliasKeys,
            string alias,
            bool allowStructuredSuffix) {
            string key = (allowStructuredSuffix ? "S:" : "N:") + alias;
            if (aliasKeys.Add(key)) {
                aliases.Add(new FormulaDependencyAlias(alias, allowStructuredSuffix));
            }
        }

        private void AddFormulaAliasDependencies(
            string formula,
            string alias,
            bool allowStructuredSuffix,
            ISet<string> dependencies) {
            foreach (string reference in FindFormulaAliasReferences(formula, alias, allowStructuredSuffix)) {
                if (TryResolveFormulaRangeReference(reference, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2)) {
                    string start = A1.CellReference(r1, c1);
                    string end = A1.CellReference(r2, c2);
                    dependencies.Add(r1 == r2 && c1 == c2
                        ? $"{sheet.Name}!{start}"
                        : $"{sheet.Name}!{start}:{end}");
                }
            }
        }

        private static IEnumerable<string> FindFormulaAliasReferences(string formula, string alias, bool allowStructuredSuffix) {
            int searchIndex = 0;
            while (searchIndex < formula.Length) {
                int index = formula.IndexOf(alias, searchIndex, StringComparison.OrdinalIgnoreCase);
                if (index < 0) {
                    yield break;
                }

                int end = index + alias.Length;
                bool validStart = index == 0
                    || (!IsFormulaAliasIdentifierCharacter(formula[index - 1]) && formula[index - 1] != '!');
                bool hasStructuredSuffix = allowStructuredSuffix && end < formula.Length && formula[end] == '[';
                if (hasStructuredSuffix) {
                    int bracketDepth = 0;
                    int cursor = end;
                    for (; cursor < formula.Length; cursor++) {
                        if (formula[cursor] == '[') {
                            bracketDepth++;
                        } else if (formula[cursor] == ']') {
                            bracketDepth--;
                            if (bracketDepth == 0) {
                                end = cursor + 1;
                                break;
                            }
                        }
                    }

                    if (bracketDepth != 0) {
                        hasStructuredSuffix = false;
                        end = index + alias.Length;
                    }
                }

                bool validEnd = end == formula.Length || !IsFormulaAliasIdentifierCharacter(formula[end]);
                if (validStart && validEnd && (!allowStructuredSuffix || hasStructuredSuffix || end == index + alias.Length)) {
                    yield return formula.Substring(index, end - index);
                }

                searchIndex = index + alias.Length;
            }
        }

        private static bool IsFormulaAliasIdentifierCharacter(char character) {
            return char.IsLetterOrDigit(character) || character == '_' || character == '.';
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
