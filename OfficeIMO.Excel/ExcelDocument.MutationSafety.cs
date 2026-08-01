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
            Action? consumeScannedElement = null,
            ExcelReference? rewriteBoundary = null,
            ExcelCellShiftDirection? cellShiftDirection = null) {
            List<Sheet> sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int editedSheetIndex = sheets.FindIndex(sheet =>
                string.Equals(sheet.Name?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase));
            if (editedSheetIndex < 0) return;
            int sr1 = 0;
            int sc1 = 0;
            int sr2 = 0;
            int sc2 = 0;
            rewriteBoundary?.GetBounds(out sr1, out sc1, out sr2, out sc2);

            foreach (MutationFormulaContext formula in EnumerateMutationFormulaContexts(sheets, editedSheetIndex)) {
                consumeScannedElement?.Invoke();
                foreach (ExcelFormulaReferenceSyntax referenceNode in ExcelFormulaSyntaxTree.Parse(formula.Text)
                    .Nodes.OfType<ExcelFormulaReferenceSyntax>()) {
                    consumeScannedElement?.Invoke();
                    if (TryGetThreeDimensionalSheetRange(
                            referenceNode.Reference,
                            out string firstSheetName,
                            out string lastSheetName)
                        && TryFindSheetIndex(sheets, firstSheetName, out int firstSheetIndex)
                        && TryFindSheetIndex(sheets, lastSheetName, out int lastSheetIndex)) {
                        int lower = Math.Min(firstSheetIndex, lastSheetIndex);
                        int upper = Math.Max(firstSheetIndex, lastSheetIndex);
                        if (editedSheetIndex >= lower && editedSheetIndex <= upper) {
                            throw new InvalidOperationException(
                                $"{operation} cannot preserve 3-D reference '{referenceNode.Text}' because worksheet '{editedSheet.Name}' is inside its sheet span.");
                        }
                    }

                    if (rewriteBoundary == null) continue;
                    ExcelReference reference = referenceNode.Reference;
                    if (!ReferenceTargetsSheet(reference, editedSheet.Name, formula.UnqualifiedTargetsEdited)) continue;
                    reference.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
                    bool intersects = rr1 <= sr2 && rr2 >= sr1 && rc1 <= sc2 && rc2 >= sc1;
                    bool contained = rr1 >= sr1 && rr2 <= sr2 && rc1 >= sc1 && rc2 <= sc2;
                    bool unsafePartial = intersects && !contained;
                    if (cellShiftDirection == ExcelCellShiftDirection.Left) {
                        unsafePartial = intersects && (rr1 < sr1 || rr2 > sr2);
                    } else if (cellShiftDirection == ExcelCellShiftDirection.Up) {
                        unsafePartial = intersects && (rc1 < sc1 || rc2 > sc2);
                    }
                    if (!unsafePartial) continue;
                    throw new InvalidOperationException(
                        $"{operation} cannot preserve partially overlapping reference '{referenceNode.Text}'. Edit the complete referenced range or update the formula first.");
                }
            }
        }

        private IEnumerable<MutationFormulaContext> EnumerateMutationFormulaContexts(
            IReadOnlyList<Sheet> sheets,
            int editedSheetIndex) {
            foreach (DefinedName name in WorkbookRoot.DefinedNames?.Elements<DefinedName>()
                ?? Enumerable.Empty<DefinedName>()) {
                if (!string.IsNullOrEmpty(name.Text)) {
                    bool local = name.LocalSheetId?.Value is uint localIndex && localIndex == (uint)editedSheetIndex;
                    yield return new MutationFormulaContext(name.Text, local);
                }
            }

            for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
                Sheet sheet = sheets[sheetIndex];
                if (sheet.Id?.Value is not string relationshipId) continue;
                OpenXmlPart part = WorkbookPartRoot.GetPartById(relationshipId);
                if (part is WorksheetPart worksheetPart) {
                    bool ownerIsEdited = sheetIndex == editedSheetIndex;
                    foreach (string formula in EnumerateMutationFormulaTexts(worksheetPart.Worksheet)) yield return new MutationFormulaContext(formula, ownerIsEdited);
                    foreach (Hyperlink hyperlink in worksheetPart.Worksheet?.Descendants<Hyperlink>()
                        ?? Enumerable.Empty<Hyperlink>()) {
                        if (string.IsNullOrWhiteSpace(hyperlink.Id?.Value)
                            && !string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) {
                            yield return new MutationFormulaContext(hyperlink.Location!.Value!, ownerIsEdited);
                        }
                    }
                    foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                        foreach (string formula in EnumerateMutationFormulaTexts(tablePart.Table)) yield return new MutationFormulaContext(formula, ownerIsEdited);
                    }
                    foreach (PivotTablePart pivotPart in worksheetPart.PivotTableParts) {
                        foreach (string formula in EnumerateMutationFormulaTexts(pivotPart.PivotTableDefinition)) yield return new MutationFormulaContext(formula, ownerIsEdited);
                    }
                    foreach (string formula in EnumerateMutationDrawingFormulaTexts(worksheetPart.DrawingsPart)) yield return new MutationFormulaContext(formula, ownerIsEdited);
                } else if (part is ChartsheetPart chartsheetPart) {
                    foreach (string formula in EnumerateMutationDrawingFormulaTexts(chartsheetPart.DrawingsPart)) yield return new MutationFormulaContext(formula, false);
                }
            }
        }

        private readonly struct MutationFormulaContext {
            internal MutationFormulaContext(string text, bool unqualifiedTargetsEdited) {
                Text = text;
                UnqualifiedTargetsEdited = unqualifiedTargetsEdited;
            }

            internal string Text { get; }
            internal bool UnqualifiedTargetsEdited { get; }
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
