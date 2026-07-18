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

        private readonly struct FormulaDependencyAliasMatch {
            internal FormulaDependencyAliasMatch(int index, FormulaDependencyAlias alias) {
                Index = index;
                Alias = alias;
            }

            internal int Index { get; }
            internal FormulaDependencyAlias Alias { get; }
        }

        private sealed class FormulaDependencyAliasCatalog {
            private sealed class TrieNode {
                internal Dictionary<char, TrieNode> Children { get; } = new Dictionary<char, TrieNode>();
                internal List<FormulaDependencyAlias> Aliases { get; } = new List<FormulaDependencyAlias>();
            }

            private readonly TrieNode _root = new TrieNode();

            internal FormulaDependencyAliasCatalog(IEnumerable<FormulaDependencyAlias> aliases) {
                foreach (FormulaDependencyAlias alias in aliases) {
                    TrieNode node = _root;
                    foreach (char character in alias.Text) {
                        char key = char.ToUpperInvariant(character);
                        if (!node.Children.TryGetValue(key, out TrieNode? child)) {
                            child = new TrieNode();
                            node.Children.Add(key, child);
                        }

                        node = child;
                    }

                    node.Aliases.Add(alias);
                }
            }

            internal IEnumerable<FormulaDependencyAliasMatch> FindMatches(string formula) {
                for (int start = 0; start < formula.Length; start++) {
                    TrieNode node = _root;
                    for (int position = start; position < formula.Length; position++) {
                        if (!node.Children.TryGetValue(char.ToUpperInvariant(formula[position]), out TrieNode? child)) {
                            break;
                        }

                        node = child;
                        foreach (FormulaDependencyAlias alias in node.Aliases) {
                            yield return new FormulaDependencyAliasMatch(start, alias);
                        }
                    }
                }
            }
        }

        private sealed class FormulaDependencyInspectionContext {
            private readonly Dictionary<string, IReadOnlyList<Cell>> _formulaCells =
                new Dictionary<string, IReadOnlyList<Cell>>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>> _sharedFormulaDefinitions =
                new Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>(StringComparer.OrdinalIgnoreCase);

            internal FormulaDependencyInspectionContext(
                ExcelSheet sourceSheet,
                IReadOnlyList<Cell> sourceFormulaCells,
                IReadOnlyDictionary<uint, SharedFormulaDefinition> sourceSharedFormulaDefinitions) {
                _formulaCells.Add(sourceSheet.Name, sourceFormulaCells);
                _sharedFormulaDefinitions.Add(sourceSheet.Name, sourceSharedFormulaDefinitions);
            }

            internal IReadOnlyList<Cell> GetFormulaCells(ExcelSheet sheet) {
                if (!_formulaCells.TryGetValue(sheet.Name, out IReadOnlyList<Cell>? cells)) {
                    cells = sheet.WorksheetRoot.Descendants<Cell>()
                        .Where(cell => cell.CellFormula != null)
                        .ToList();
                    _formulaCells.Add(sheet.Name, cells);
                }

                return cells;
            }

            internal IReadOnlyDictionary<uint, SharedFormulaDefinition> GetSharedFormulaDefinitions(ExcelSheet sheet) {
                if (!_sharedFormulaDefinitions.TryGetValue(
                    sheet.Name,
                    out IReadOnlyDictionary<uint, SharedFormulaDefinition>? definitions)) {
                    definitions = sheet.BuildSharedFormulaDefinitions();
                    _sharedFormulaDefinitions.Add(sheet.Name, definitions);
                }

                return definitions;
            }
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
            string? sourceCellReference,
            string formula,
            FormulaDependencyAliasCatalog aliases,
            FormulaDependencyTableCatalog tables) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return Array.Empty<string>();
            }

            try {
                bool hasSourceCell = TryParseCellReference(
                    sourceCellReference ?? string.Empty,
                    out int parsedSourceRow,
                    out int parsedSourceColumn);
                int? sourceRow = hasSourceCell ? parsedSourceRow : (int?)null;
                string searchableFormula = MaskFormulaStringLiterals(formula);
                int? sourceColumn = hasSourceCell ? parsedSourceColumn : (int?)null;
                string valueDependencyFormula = MaskFormulaReferenceShapeArguments(
                    searchableFormula,
                    sourceRow,
                    sourceColumn);
                string localReferenceFormula = MaskFormulaNonLocalReferenceSegments(valueDependencyFormula);
                string directReferenceFormula = MaskFormulaStructuredReferenceSegments(localReferenceFormula);
                List<FormulaDependencyReferenceMatch> dependencyMatches = FormulaReferenceRegex.Matches(directReferenceFormula)
                    .Cast<Match>()
                    .Where(match => IsLocalFormulaReferenceMatch(searchableFormula, match))
                    .Where(match => !IsFormulaDependencyFunctionToken(searchableFormula, match))
                    .Where(match => !string.IsNullOrWhiteSpace(match.Groups["reference"].Value))
                    .Select(match => new FormulaDependencyReferenceMatch(
                        match.Index,
                        match.Length,
                        match.Groups["reference"].Value))
                    .ToList();
                var dependencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                IReadOnlyList<FormulaLexicalBinding> lexicalBindings = GetFormulaLexicalBindings(searchableFormula);
                foreach (FormulaDependencyAliasMatch match in aliases.FindMatches(valueDependencyFormula)) {
                    AddFormulaAliasDependencyMatch(valueDependencyFormula, match, lexicalBindings, sourceRow, dependencyMatches);
                }
                if (hasSourceCell) {
                    AddUnqualifiedCurrentRowDependencyMatches(
                        localReferenceFormula,
                        parsedSourceRow,
                        parsedSourceColumn,
                        tables,
                        dependencyMatches);
                }
                AddFormulaDependencies(searchableFormula, dependencyMatches, sourceRow, dependencies);

                return dependencies
                    .OrderBy(reference => reference, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            } catch (RegexMatchTimeoutException) {
                return Array.Empty<string>();
            }
        }

        private FormulaDependencyAliasCatalog GetFormulaDependencyAliases() {
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

            return new FormulaDependencyAliasCatalog(aliases);
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

        private void AddFormulaAliasDependencyMatch(
            string formula,
            FormulaDependencyAliasMatch match,
            IReadOnlyList<FormulaLexicalBinding> lexicalBindings,
            int? sourceRow,
            ICollection<FormulaDependencyReferenceMatch> dependencyMatches) {
            if (!TryGetFormulaAliasReference(formula, match, out string reference)) {
                return;
            }

            bool hasStructuredSuffix = match.Alias.AllowStructuredSuffix
                && reference.Length > match.Alias.Text.Length;
            if ((!hasStructuredSuffix
                    && lexicalBindings.Any(binding => binding.Shadows(match.Alias.Text, match.Index, match.Alias.Text.Length)))
                || !TryResolveFormulaRangeReference(reference, sourceRow, out _, out _, out _, out _, out _)) {
                return;
            }

            dependencyMatches.Add(new FormulaDependencyReferenceMatch(match.Index, reference.Length, reference));
        }

        private static bool TryGetFormulaAliasReference(
            string formula,
            FormulaDependencyAliasMatch match,
            out string reference) {
            reference = string.Empty;
            string alias = match.Alias.Text;
            int index = match.Index;
            int end = index + alias.Length;
            bool validStart = index == 0
                || (!IsFormulaAliasIdentifierCharacter(formula[index - 1])
                    && formula[index - 1] != '!'
                    && formula[index - 1] != ']');
            if (!validStart
                || IsInsideFormulaErrorLiteral(formula, index)
                || IsInsideFormulaStructuredReference(formula, index)
                || IsInsideQuotedFormulaSheetQualifier(formula, index)) {
                return false;
            }

            bool hasStructuredSuffix = match.Alias.AllowStructuredSuffix
                && end < formula.Length
                && formula[end] == '[';
            if (!match.Alias.AllowStructuredSuffix && end < formula.Length && formula[end] == '[') {
                return false;
            }

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
                    return false;
                }
            }

            if (end < formula.Length && IsFormulaAliasIdentifierCharacter(formula[end])) {
                return false;
            }

            int nextToken = end;
            while (nextToken < formula.Length && char.IsWhiteSpace(formula[nextToken])) {
                nextToken++;
            }

            if (nextToken < formula.Length
                && (formula[nextToken] == '('
                    || formula[nextToken] == '!'
                    || IsStartOfThreeDimensionalReference(formula, nextToken))) {
                return false;
            }

            reference = formula.Substring(index, end - index);
            return true;
        }

        private static bool IsInsideFormulaErrorLiteral(string formula, int index) {
            int start = index;
            while (start > 0 && IsFormulaErrorLiteralCharacter(formula[start - 1])) {
                start--;
            }

            int end = index;
            while (end < formula.Length && IsFormulaErrorLiteralCharacter(formula[end])) {
                end++;
            }

            return start < end
                && formula[start] == '#'
                && TryParseFormulaErrorLiteral(formula.Substring(start, end - start), out _);
        }

        private static bool IsFormulaErrorLiteralCharacter(char character) {
            return char.IsLetterOrDigit(character)
                || character == '#'
                || character == '/'
                || character == '!'
                || character == '?';
        }

        private static bool IsInsideFormulaStructuredReference(string formula, int index) {
            int bracketDepth = 0;
            for (int position = 0; position < index; position++) {
                if (formula[position] == '[') {
                    bracketDepth++;
                } else if (formula[position] == ']' && bracketDepth > 0) {
                    bracketDepth--;
                }
            }

            return bracketDepth > 0;
        }

        private static bool IsFormulaAliasIdentifierCharacter(char character) {
            return char.IsLetterOrDigit(character) || character == '_' || character == '.' || character == '\\';
        }

        private static bool IsLocalFormulaReferenceMatch(string formula, Match match) {
            if (!match.Success || match.Index < 0 || match.Index + match.Length > formula.Length) {
                return false;
            }

            string originalToken = formula.Substring(match.Index, match.Length);
            if (originalToken.IndexOf('[') >= 0 || originalToken.IndexOf(']') >= 0) {
                return false;
            }

            if (match.Index > 0 && formula[match.Index - 1] == ']') {
                return false;
            }

            string reference = match.Groups["reference"].Value;
            int qualifierSeparator = reference.LastIndexOf('!');
            if (qualifierSeparator > 0) {
                string qualifier = reference.Substring(0, qualifierSeparator);
                if (qualifier.IndexOf(':') >= 0 || qualifier.IndexOf('[') >= 0 || qualifier.IndexOf(']') >= 0) {
                    return false;
                }
            }

            return true;
        }

        private bool IsFormulaDependencyFunctionToken(string formula, Match match) {
            string token = match.Groups["reference"].Value;
            if (token.IndexOf('!') >= 0 || token.IndexOf(':') >= 0 || token.IndexOf('$') >= 0) {
                return false;
            }

            int cursor = match.Index + match.Length;
            int whitespaceStart = cursor;
            while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }

            if (cursor >= formula.Length || formula[cursor] != '(') {
                return false;
            }

            return cursor == whitespaceStart
                || ExcelFormulaCapabilities.IsBuiltInFunction(token)
                || _excelDocument.Calculation.TryGetCustomFunction(token, out _);
        }

        private static bool IsStartOfThreeDimensionalReference(string formula, int index) {
            if (index >= formula.Length || formula[index] != ':') {
                return false;
            }

            for (int cursor = index + 1; cursor < formula.Length; cursor++) {
                char character = formula[cursor];
                if (character == '!') {
                    return true;
                }
                if (character == '+'
                    || character == '-'
                    || character == '*'
                    || character == '/'
                    || character == '^'
                    || character == '&'
                    || character == '='
                    || character == '<'
                    || character == '>'
                    || character == ','
                    || character == ';'
                    || character == '('
                    || character == ')') {
                    return false;
                }
            }

            return false;
        }

        private IReadOnlyList<string> GetFormulaDependencyIssues(
            string? sourceCellReference,
            IReadOnlyList<string> dependencies,
            FormulaDependencyInspectionContext inspectionContext) {
            if (dependencies.Count == 0) {
                return Array.Empty<string>();
            }

            var issues = new List<string>();
            string? sourceReference = NormalizeFormulaCellReference(sourceCellReference);
            foreach (string dependency in dependencies) {
                if (!TryResolveFormulaDependencyReference(
                    dependency,
                    out ExcelSheet dependencySheet,
                    out int r1,
                    out int c1,
                    out int r2,
                    out int c2,
                    out _)) {
                    issues.Add($"Cannot resolve dependency '{dependency}'.");
                    continue;
                }

                if (sourceReference != null
                    && string.Equals(dependencySheet.Name, Name, StringComparison.OrdinalIgnoreCase)
                    && TryParseCellReference(sourceReference, out int sourceRow, out int sourceColumn)
                    && sourceRow >= r1 && sourceRow <= r2 && sourceColumn >= c1 && sourceColumn <= c2) {
                    issues.Add($"Dependency '{dependency}' references its own formula cell.");
                }

                foreach (Cell dependencyCell in inspectionContext.GetFormulaCells(dependencySheet)) {
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

                    if (!dependencySheet.TryEvaluateFormulaCellValue(
                        dependencyCell,
                        out _,
                        inspectionContext.GetSharedFormulaDefinitions(dependencySheet))) {
                        issues.Add($"Dependency '{formattedDependencyCell}' contains a formula outside the lightweight evaluator support.");
                    }
                }
            }

            return issues
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(issue => issue, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string NormalizeFormulaDependencyReference(string reference, int? sourceRow) {
            string normalized = reference.Trim().Replace("$", string.Empty);
            if (TryResolveFormulaDependencyReference(
                normalized,
                sourceRow,
                out ExcelSheet sheet,
                out _,
                out _,
                out _,
                out _,
                out string address)) {
                return $"{sheet.Name}!{address}";
            }

            return normalized;
        }

        private bool TryResolveFormulaDependencyReference(
            string token,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2,
            out string address) {
            return TryResolveFormulaDependencyReference(token, null, out sheet, out r1, out c1, out r2, out c2, out address);
        }

        private bool TryResolveFormulaDependencyReference(
            string token,
            int? sourceRow,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2,
            out string address) {
            if (TryResolveFormulaRangeReference(token, sourceRow, out sheet, out r1, out c1, out r2, out c2)) {
                string start = A1.CellReference(r1, c1);
                string end = A1.CellReference(r2, c2);
                address = r1 == r2 && c1 == c2 ? start : start + ":" + end;
                return true;
            }

            sheet = this;
            r1 = c1 = r2 = c2 = 0;
            address = string.Empty;
            return TryParseQualifiedFormulaWholeRange(
                token,
                null,
                out sheet,
                out r1,
                out c1,
                out r2,
                out c2,
                out address);
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

        private static bool IsInsideQuotedFormulaSheetQualifier(string formula, int index) {
            for (int start = 0; start < formula.Length; start++) {
                if (formula[start] != '\'') {
                    continue;
                }

                int cursor = start + 1;
                while (cursor < formula.Length) {
                    if (formula[cursor] != '\'') {
                        cursor++;
                        continue;
                    }
                    if (cursor + 1 < formula.Length && formula[cursor + 1] == '\'') {
                        cursor += 2;
                        continue;
                    }

                    break;
                }

                if (cursor >= formula.Length) {
                    return false;
                }
                if (cursor + 1 >= formula.Length || formula[cursor + 1] != '!') {
                    continue;
                }

                if (index > start && index <= cursor) {
                    return true;
                }
                start = cursor;
            }

            return false;
        }

        private static string MaskFormulaStructuredReferenceSegments(string formula) {
            var builder = new StringBuilder(formula.Length);
            int bracketDepth = 0;
            foreach (char character in formula) {
                if (character == '[') {
                    bracketDepth++;
                    builder.Append(' ');
                } else if (character == ']' && bracketDepth > 0) {
                    bracketDepth--;
                    builder.Append(' ');
                } else {
                    builder.Append(bracketDepth > 0 ? ' ' : character);
                }
            }

            return builder.ToString();
        }
    }
}
