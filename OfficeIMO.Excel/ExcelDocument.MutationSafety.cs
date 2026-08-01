using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Rejects structural edits whose package-wide references cannot be preserved safely.</summary>
        internal void ValidateMutationReferencesCanBeRewritten(
            ExcelSheet editedSheet,
            string operation,
            Action? consumeScannedElement = null) {
            List<Sheet> sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int editedSheetIndex = sheets.FindIndex(sheet =>
                string.Equals(sheet.Name?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase));
            if (editedSheetIndex < 0) return;

            foreach (string formula in EnumerateMutationFormulaTexts(sheets)) {
                consumeScannedElement?.Invoke();
                foreach (ExcelFormulaReferenceSyntax referenceNode in ExcelFormulaSyntaxTree.Parse(formula)
                    .Nodes.OfType<ExcelFormulaReferenceSyntax>()) {
                    consumeScannedElement?.Invoke();
                    if (!TryGetThreeDimensionalSheetRange(
                            referenceNode.Reference,
                            out string firstSheetName,
                            out string lastSheetName)
                        || !TryFindSheetIndex(sheets, firstSheetName, out int firstSheetIndex)
                        || !TryFindSheetIndex(sheets, lastSheetName, out int lastSheetIndex)) continue;

                    int lower = Math.Min(firstSheetIndex, lastSheetIndex);
                    int upper = Math.Max(firstSheetIndex, lastSheetIndex);
                    if (editedSheetIndex < lower || editedSheetIndex > upper) continue;
                    throw new InvalidOperationException(
                        $"{operation} cannot preserve 3-D reference '{referenceNode.Text}' because worksheet '{editedSheet.Name}' is inside its sheet span.");
                }
            }
        }

        private IEnumerable<string> EnumerateMutationFormulaTexts(IReadOnlyList<Sheet> sheets) {
            foreach (DefinedName name in WorkbookRoot.DefinedNames?.Elements<DefinedName>()
                ?? Enumerable.Empty<DefinedName>()) {
                if (!string.IsNullOrEmpty(name.Text)) yield return name.Text;
            }

            foreach (Sheet sheet in sheets) {
                if (sheet.Id?.Value is not string relationshipId) continue;
                OpenXmlPart part = WorkbookPartRoot.GetPartById(relationshipId);
                if (part is WorksheetPart worksheetPart) {
                    foreach (string formula in EnumerateMutationFormulaTexts(worksheetPart.Worksheet)) yield return formula;
                    foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                        foreach (string formula in EnumerateMutationFormulaTexts(tablePart.Table)) yield return formula;
                    }
                    foreach (PivotTablePart pivotPart in worksheetPart.PivotTableParts) {
                        foreach (string formula in EnumerateMutationFormulaTexts(pivotPart.PivotTableDefinition)) yield return formula;
                    }
                    foreach (string formula in EnumerateMutationDrawingFormulaTexts(worksheetPart.DrawingsPart)) yield return formula;
                } else if (part is ChartsheetPart chartsheetPart) {
                    foreach (string formula in EnumerateMutationDrawingFormulaTexts(chartsheetPart.DrawingsPart)) yield return formula;
                }
            }
        }

        private static IEnumerable<string> EnumerateMutationDrawingFormulaTexts(DrawingsPart? drawingsPart) {
            if (drawingsPart == null) yield break;
            foreach (Xdr.Shape shape in drawingsPart.WorksheetDrawing?.Descendants<Xdr.Shape>()
                ?? Enumerable.Empty<Xdr.Shape>()) {
                if (!string.IsNullOrEmpty(shape.TextLink?.Value)) yield return shape.TextLink!.Value!;
            }
            foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                foreach (string formula in EnumerateMutationFormulaTexts(chartPart.ChartSpace)) yield return formula;
            }
            foreach (ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                foreach (string formula in EnumerateMutationFormulaTexts(chartPart.ChartSpace)) yield return formula;
            }
        }

        private static IEnumerable<string> EnumerateMutationFormulaTexts(OpenXmlPartRootElement? root) {
            if (root == null) yield break;
            foreach (OpenXmlLeafTextElement leaf in root.Descendants<OpenXmlLeafTextElement>()
                .Where(IsMutationFormulaLeaf)) {
                if (!string.IsNullOrEmpty(leaf.Text)) yield return leaf.Text;
            }
        }

        private static bool TryGetThreeDimensionalSheetRange(
            ExcelReference reference,
            out string firstSheetName,
            out string lastSheetName) {
            firstSheetName = string.Empty;
            lastSheetName = string.Empty;
            string qualifier = reference.Qualifier?.Trim() ?? string.Empty;
            if (qualifier.Length == 0) return false;

            int quotedSeparator = qualifier.IndexOf("':'", StringComparison.Ordinal);
            if (quotedSeparator >= 0) {
                firstSheetName = UnquoteMutationSheetToken(qualifier.Substring(0, quotedSeparator + 1));
                lastSheetName = UnquoteMutationSheetToken(qualifier.Substring(quotedSeparator + 2));
            } else {
                string normalized = UnquoteMutationSheetToken(qualifier);
                int separator = normalized.IndexOf(':');
                if (separator <= 0 || separator != normalized.LastIndexOf(':')) return false;
                firstSheetName = normalized.Substring(0, separator);
                lastSheetName = normalized.Substring(separator + 1);
            }

            return firstSheetName.Length > 0
                && lastSheetName.Length > 0
                && !firstSheetName.StartsWith("[", StringComparison.Ordinal)
                && !lastSheetName.StartsWith("[", StringComparison.Ordinal);
        }

        private static string UnquoteMutationSheetToken(string value) {
            string result = value.Trim();
            if (result.Length >= 2 && result[0] == '\'' && result[result.Length - 1] == '\'') {
                result = result.Substring(1, result.Length - 2).Replace("''", "'");
            }
            return result;
        }

        private static bool TryFindSheetIndex(
            IReadOnlyList<Sheet> sheets,
            string sheetName,
            out int sheetIndex) {
            for (int index = 0; index < sheets.Count; index++) {
                if (SheetNameLookup.Matches(sheets[index].Name?.Value, sheetName)) {
                    sheetIndex = index;
                    return true;
                }
            }
            sheetIndex = -1;
            return false;
        }
    }
}
