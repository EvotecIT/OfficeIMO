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
            IReadOnlyDictionary<int, int>? externalReferenceMap = null,
            DefinedNameCopyBudget? budget = null) {
            budget ??= new DefinedNameCopyBudget(4096, 1_000_000);
            DefinedNames? sourceDefinedNames = sourceDocument.WorkbookRoot.DefinedNames;
            if (sourceDefinedNames == null) {
                return;
            }

            ushort targetSheetPosition = GetSheetPositionIndex(targetSheet);
            WorksheetPart targetWorksheetPart = targetSheet.WorksheetPart;
            List<string> formulaTexts = CollectFormulaTexts(targetWorksheetPart).ToList();
            if (formulaTexts.Count == 0) {
                return;
            }

            var sourceSheetNamesByPosition = sourceDocument.GetSheetNamesByPosition();
            string? currentSourceSheetName = sheetNameMap
                .FirstOrDefault(mapping => string.Equals(mapping.Value, targetSheet.Name, StringComparison.OrdinalIgnoreCase))
                .Key;
            var copiedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<DefinedName> sourceNames = sourceDefinedNames.Elements<DefinedName>()
                .Where(name => !string.IsNullOrWhiteSpace(name.Name?.Value)
                    && !name.Name!.Value!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var localNamesForCurrentSourceSheet = new HashSet<string>(sourceNames
                .Where(name => name.LocalSheetId != null
                    && !string.IsNullOrEmpty(currentSourceSheetName)
                    && sourceSheetNamesByPosition.TryGetValue((ushort)name.LocalSheetId.Value, out string? owner)
                    && string.Equals(owner, currentSourceSheetName, StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(name.Name?.Value))
                .Select(name => name.Name!.Value!), StringComparer.OrdinalIgnoreCase);
            var sourceNamesByName = sourceNames
                .GroupBy(name => name.Name!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
            var pendingNames = new Queue<string>();
            var queuedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string formulaText in formulaTexts) {
                EnqueueReferencedDefinedNames(formulaText, sourceNamesByName, queuedNames, pendingNames);
            }

            var plannedCopies = new List<(DefinedName Clone, ushort DestinationSheetPosition, string Name)>();
            while (pendingNames.Count > 0) {
                string referencedName = pendingNames.Dequeue();
                foreach (DefinedName sourceName in sourceNamesByName[referencedName]) {
                    string name = sourceName.Name!.Value!;
                    if (sourceName.LocalSheetId == null && localNamesForCurrentSourceSheet.Contains(name)) {
                        continue;
                    }

                    string copyKey = name + "|" + (sourceName.LocalSheetId?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    if (!copiedNames.Add(copyKey)) {
                        continue;
                    }

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

                        EnqueueReferencedDefinedNames(clone.Text!, sourceNamesByName, queuedNames, pendingNames);
                    }

                    budget.Consume(clone.Text?.Length ?? 0);
                    plannedCopies.Add((clone, destinationSheetPosition, name));
                }
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

        private static void EnqueueReferencedDefinedNames(
            string formula,
            IReadOnlyDictionary<string, List<DefinedName>> sourceNamesByName,
            ISet<string> queuedNames,
            Queue<string> pendingNames) {
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
                if (sourceNamesByName.ContainsKey(token) && queuedNames.Add(token)) {
                    pendingNames.Enqueue(token);
                }
            }
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
