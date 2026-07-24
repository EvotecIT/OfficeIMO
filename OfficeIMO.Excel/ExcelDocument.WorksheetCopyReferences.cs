using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void RewriteCopiedWorksheetReferences(
            WorksheetPart worksheetPart,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null) {
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            bool worksheetChanged = RewriteWorksheetSheetReferences(worksheet, sheetNameMap);
            if (tableNameMap?.Count > 0) {
                RewriteStructuredTableReferences(worksheet, tableNameMap);
                worksheetChanged = true;
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                Table? table = tablePart.Table;
                if (table == null) {
                    continue;
                }

                bool tableChanged = false;
                foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>()) {
                    tableChanged |= RewriteFormulaSheetReference(formula, sheetNameMap);
                }

                foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>()) {
                    tableChanged |= RewriteFormulaSheetReference(formula, sheetNameMap);
                }

                if (tableNameMap?.Count > 0) {
                    RewriteStructuredTableReferences(table, tableNameMap);
                    tableChanged = true;
                }

                if (tableChanged) {
                    table.Save();
                }
            }

            if (worksheetChanged) {
                worksheet.Save();
            }
        }

        private void CopyReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            ExcelSheet targetSheet,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            IReadOnlyDictionary<int, int>? externalReferenceMap = null) {
            ushort targetSheetPosition = GetSheetPositionIndex(targetSheet);
            var sourceSheetNamesByPosition = sourceDocument.GetSheetNamesByPosition();
            string? currentSourceSheetName = sheetNameMap
                .FirstOrDefault(mapping => string.Equals(mapping.Value, targetSheet.Name, StringComparison.OrdinalIgnoreCase))
                .Key;
            if (string.IsNullOrEmpty(currentSourceSheetName)) {
                return;
            }

            ExcelSheet sourceSheet = sourceDocument.GetSheet(currentSourceSheetName!);
            var copiedSourceSheetNames = new HashSet<string>(sheetNameMap.Keys, StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<DefinedName> referencedSourceNames = ResolveReferencedDefinedNamesFromSource(
                sourceDocument,
                sourceSheet,
                copiedSourceSheetNames);
            var plannedCopies = new List<(DefinedName Clone, ushort DestinationSheetPosition, string Name)>();
            foreach (DefinedName sourceName in referencedSourceNames) {
                string name = sourceName.Name!.Value!;
                ushort destinationSheetPosition = targetSheetPosition;
                if (sourceName.LocalSheetId != null) {
                    if (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? sourceNameOwner)
                        || !sheetNameMap.TryGetValue(sourceNameOwner, out string? targetNameOwner)
                        || !TryGetSheetPositionIndexByName(targetNameOwner, out destinationSheetPosition)) {
                        continue;
                    }
                }

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.LocalSheetId = destinationSheetPosition;
                clone.Name = name;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ReplaceSheetNameReferences(clone.Text!, sheetNameMap);
                    if (tableNameMap?.Count > 0) {
                        clone.Text = RewriteStructuredTableReferences(clone.Text!, tableNameMap);
                    }

                    if (externalReferenceMap?.Count > 0) {
                        clone.Text = RewriteExternalWorkbookReferenceIndexes(clone.Text!, externalReferenceMap);
                    }
                }

                plannedCopies.Add((clone, destinationSheetPosition, name));
            }

            if (plannedCopies.Count == 0) {
                return;
            }

            DefinedNames targetDefinedNames = WorkbookRoot.DefinedNames ??= new DefinedNames();
            foreach (var plannedCopy in plannedCopies) {
                foreach (DefinedName existing in targetDefinedNames.Elements<DefinedName>()
                    .Where(item => item.LocalSheetId != null
                        && item.LocalSheetId.Value == plannedCopy.DestinationSheetPosition
                        && string.Equals(item.Name?.Value, plannedCopy.Name, StringComparison.OrdinalIgnoreCase))
                    .ToList()) {
                    existing.Remove();
                }

                targetDefinedNames.Append(plannedCopy.Clone);
            }

            WorkbookRoot.Save();
        }

        private static void PreflightReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            IReadOnlyList<ExcelSheet> sourceSheets,
            DefinedNameCopyBudget budget) {
            var copiedSourceSheetNames = new HashSet<string>(
                sourceSheets.Select(sheet => sheet.Name),
                StringComparer.OrdinalIgnoreCase);
            foreach (ExcelSheet sourceSheet in sourceSheets) {
                IReadOnlyList<DefinedName> referencedNames = ResolveReferencedDefinedNamesFromSource(
                    sourceDocument,
                    sourceSheet,
                    copiedSourceSheetNames);
                foreach (DefinedName sourceName in referencedNames) {
                    budget.Consume(sourceName.Text?.Length ?? 0);
                }
            }
        }

        private static IReadOnlyList<DefinedName> ResolveReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            ExcelSheet sourceSheet,
            ISet<string> copiedSourceSheetNames) {
            DefinedNames? sourceDefinedNames = sourceDocument.WorkbookRoot.DefinedNames;
            if (sourceDefinedNames == null) {
                return Array.Empty<DefinedName>();
            }

            List<string> formulaTexts = CollectFormulaTexts(sourceSheet.WorksheetPart).ToList();
            if (formulaTexts.Count == 0) {
                return Array.Empty<DefinedName>();
            }

            var sourceSheetNamesByPosition = sourceDocument.GetSheetNamesByPosition();
            List<DefinedName> sourceNames = sourceDefinedNames.Elements<DefinedName>()
                .Where(name => !string.IsNullOrWhiteSpace(name.Name?.Value)
                    && !name.Name!.Value!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var sourceNamesByName = sourceNames
                .GroupBy(name => name.Name!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
            var pendingNames = new Queue<DefinedNameReference>();
            var queuedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string formulaText in formulaTexts) {
                EnqueueReferencedDefinedNames(
                    formulaText,
                    sourceSheet.Name,
                    sourceNamesByName,
                    queuedNames,
                    pendingNames);
            }

            var resolvedNames = new List<DefinedName>();
            var copiedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingNames.Count > 0) {
                DefinedNameReference reference = pendingNames.Dequeue();
                List<DefinedName> candidates = sourceNamesByName[reference.Name];
                bool hasLocalNameInContext = candidates.Any(candidate =>
                    candidate.LocalSheetId != null
                    && sourceSheetNamesByPosition.TryGetValue((ushort)candidate.LocalSheetId.Value, out string? candidateOwner)
                    && string.Equals(candidateOwner, reference.ContextSheetName, StringComparison.OrdinalIgnoreCase));
                foreach (DefinedName sourceName in candidates) {
                    string name = sourceName.Name!.Value!;
                    string dependencyContext = reference.ContextSheetName;
                    if (sourceName.LocalSheetId == null) {
                        if (reference.ExplicitSheetName != null || hasLocalNameInContext) {
                            continue;
                        }
                    } else {
                        if (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? localOwner)) {
                            continue;
                        }

                        string requiredOwner = reference.ExplicitSheetName ?? reference.ContextSheetName;
                        if (!string.Equals(localOwner, requiredOwner, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }

                        dependencyContext = localOwner;
                    }

                    string copyKey = name + "|" + (sourceName.LocalSheetId?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    if (!copiedNames.Add(copyKey)) {
                        continue;
                    }

                    if (sourceName.LocalSheetId != null
                        && (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? copiedOwner)
                            || !copiedSourceSheetNames.Contains(copiedOwner))) {
                        continue;
                    }

                    resolvedNames.Add(sourceName);
                    if (!string.IsNullOrEmpty(sourceName.Text)) {
                        EnqueueReferencedDefinedNames(
                            sourceName.Text!,
                            dependencyContext,
                            sourceNamesByName,
                            queuedNames,
                            pendingNames);
                    }
                }
            }

            return resolvedNames;
        }

        private static void EnqueueReferencedDefinedNames(
            string formula,
            string contextSheetName,
            IReadOnlyDictionary<string, List<DefinedName>> sourceNamesByName,
            ISet<string> queuedNames,
            Queue<DefinedNameReference> pendingNames) {
            int tokenStart = -1;
            for (int index = 0; index <= formula.Length; index++) {
                bool isTokenCharacter = index < formula.Length && IsDefinedNameCharacter(formula[index]);
                if (isTokenCharacter) {
                    if (tokenStart < 0) {
                        tokenStart = index;
                    }

                    continue;
                }

                if (tokenStart < 0) {
                    continue;
                }

                string token = formula.Substring(tokenStart, index - tokenStart);
                tokenStart = -1;
                if (!sourceNamesByName.ContainsKey(token)) {
                    continue;
                }

                string? explicitSheetName = TryReadSheetQualifier(formula, index - token.Length);
                var reference = new DefinedNameReference(token, contextSheetName, explicitSheetName);
                if (queuedNames.Add(reference.Key)) {
                    pendingNames.Enqueue(reference);
                }
            }
        }

        private static string? TryReadSheetQualifier(string formula, int tokenStart) {
            int bangIndex = tokenStart - 1;
            if (bangIndex < 0 || formula[bangIndex] != '!') {
                return null;
            }

            int qualifierEnd = bangIndex - 1;
            if (qualifierEnd < 0) {
                return null;
            }

            if (formula[qualifierEnd] == '\'') {
                for (int index = qualifierEnd - 1; index >= 0; index--) {
                    if (formula[index] != '\'') {
                        continue;
                    }

                    if (index > 0 && formula[index - 1] == '\'') {
                        index--;
                        continue;
                    }

                    return formula.Substring(index + 1, qualifierEnd - index - 1).Replace("''", "'");
                }

                return null;
            }

            int qualifierStart = qualifierEnd;
            while (qualifierStart >= 0 && IsDefinedNameCharacter(formula[qualifierStart])) {
                qualifierStart--;
            }

            return qualifierStart == qualifierEnd
                ? null
                : formula.Substring(qualifierStart + 1, qualifierEnd - qualifierStart);
        }

        private readonly struct DefinedNameReference {
            internal DefinedNameReference(string name, string contextSheetName, string? explicitSheetName) {
                Name = name;
                ContextSheetName = contextSheetName;
                ExplicitSheetName = explicitSheetName;
            }

            internal string Name { get; }

            internal string ContextSheetName { get; }

            internal string? ExplicitSheetName { get; }

            internal string Key => ContextSheetName + "|" + (ExplicitSheetName ?? string.Empty) + "|" + Name;
        }

        private sealed class DefinedNameCopyBudget {
            private readonly int _maximumNames;
            private readonly int _maximumCharacters;
            private int _copiedNames;
            private long _copiedCharacters;

            internal DefinedNameCopyBudget(int maximumNames, int maximumCharacters) {
                _maximumNames = maximumNames;
                _maximumCharacters = maximumCharacters;
            }

            internal void Consume(int characters) {
                if (_copiedNames >= _maximumNames) {
                    throw new InvalidOperationException(
                        $"Worksheet copy exceeds the configured defined-name limit of {_maximumNames}.");
                }

                long totalCharacters = _copiedCharacters + characters;
                if (totalCharacters > _maximumCharacters) {
                    throw new InvalidOperationException(
                        $"Worksheet copy exceeds the configured defined-name character limit of {_maximumCharacters}.");
                }

                _copiedNames++;
                _copiedCharacters = totalCharacters;
            }
        }

        private Dictionary<ushort, string> GetSheetNamesByPosition() {
            var names = new Dictionary<ushort, string>();
            ushort position = 0;
            foreach (Sheet sheet in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                string? name = sheet.Name?.Value;
                if (!string.IsNullOrEmpty(name)) {
                    names[position] = name!;
                }

                position++;
            }

            return names;
        }

        private bool TryGetSheetPositionIndexByName(string sheetName, out ushort position) {
            position = 0;
            foreach (Sheet sheet in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }

                position++;
            }

            return false;
        }

        private static IEnumerable<string> CollectFormulaTexts(WorksheetPart worksheetPart) {
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            foreach (CellFormula formula in worksheet.Descendants<CellFormula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula formula in worksheet.Descendants<Formula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula1 formula in worksheet.Descendants<Formula1>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula2 formula in worksheet.Descendants<Formula2>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (OfficeFormula formula in worksheet.Descendants<OfficeFormula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                Table? table = tablePart.Table;
                if (table == null) {
                    continue;
                }

                foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>()) {
                    if (!string.IsNullOrEmpty(formula.Text)) {
                        yield return formula.Text!;
                    }
                }

                foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>()) {
                    if (!string.IsNullOrEmpty(formula.Text)) {
                        yield return formula.Text!;
                    }
                }
            }
        }

        private static bool ContainsDefinedNameToken(string formula, string definedName) {
            for (int index = 0; index <= formula.Length - definedName.Length; index++) {
                if (!string.Equals(formula.Substring(index, definedName.Length), definedName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                int afterIndex = index + definedName.Length;
                if ((index == 0 || !IsDefinedNameCharacter(formula[index - 1]))
                    && (afterIndex == formula.Length || !IsDefinedNameCharacter(formula[afterIndex]))) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsDefinedNameCharacter(char value) {
            return char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == '\\';
        }
    }
}
