using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void RewriteCopiedWorksheetReferences(WorksheetPart worksheetPart, IReadOnlyDictionary<string, string> sheetNameMap) {
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            bool worksheetChanged = RewriteWorksheetSheetReferences(worksheet, sheetNameMap);

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
            IReadOnlyDictionary<string, string> sheetNameMap) {
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
            DefinedNames targetDefinedNames = WorkbookRoot.DefinedNames ??= new DefinedNames();
            foreach (DefinedName sourceName in sourceDefinedNames.Elements<DefinedName>().ToList()) {
                string? name = sourceName.Name?.Value;
                if (string.IsNullOrWhiteSpace(name)
                    || name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
                    || !formulaTexts.Any(text => ContainsDefinedNameToken(text, name))) {
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

                foreach (DefinedName existing in targetDefinedNames.Elements<DefinedName>()
                    .Where(item => item.LocalSheetId != null
                        && item.LocalSheetId.Value == destinationSheetPosition
                        && string.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                    .ToList()) {
                    existing.Remove();
                }

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.LocalSheetId = destinationSheetPosition;
                clone.Name = name;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ReplaceSheetNameReferences(clone.Text!, sheetNameMap);
                }

                targetDefinedNames.Append(clone);
            }

            WorkbookRoot.Save();
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
