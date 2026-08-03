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
            ExcelCellShiftDirection? cellShiftDirection = null,
            Func<ExcelReference, ExcelReference?>? capacityTransform = null) {
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
                    if (ExcelReference.TryGetThreeDimensionalSheetRange(
                            referenceNode.Reference.Qualifier,
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

                    ExcelReference reference = referenceNode.Reference;
                    if (!ReferenceTargetsSheet(reference, editedSheet.Name, formula.UnqualifiedTargetsEdited)) continue;
                    ValidateMutationReferenceCapacity(reference, referenceNode.Text, operation, capacityTransform);
                    if (rewriteBoundary == null) continue;
                    if (!IsUnsafePartialMutationReference(reference, sr1, sc1, sr2, sc2, cellShiftDirection)) continue;
                    throw new InvalidOperationException(
                        $"{operation} cannot preserve partially overlapping reference '{referenceNode.Text}'. Edit the complete referenced range or update the formula first.");
                }
            }

            if (rewriteBoundary == null && capacityTransform == null) return;
            foreach (string referenceText in EnumerateMutationRangeMetadataReferences(editedSheet.WorksheetPart)) {
                consumeScannedElement?.Invoke();
                if (!ExcelReference.TryParse(referenceText, out ExcelReference? reference)) continue;
                ValidateMutationReferenceCapacity(reference!, referenceText, operation, capacityTransform);
                if (rewriteBoundary == null
                    || !IsUnsafePartialMutationReference(reference!, sr1, sc1, sr2, sc2, cellShiftDirection)) continue;
                throw new InvalidOperationException(
                    $"{operation} cannot preserve partially overlapping range metadata '{referenceText}'. Edit the complete metadata range first.");
            }
            foreach (string referenceText in EnumerateMutationExternalRangeMetadataReferences(editedSheet)) {
                consumeScannedElement?.Invoke();
                if (!ExcelReference.TryParse(referenceText, out ExcelReference? reference)) continue;
                ValidateMutationReferenceCapacity(reference!, referenceText, operation, capacityTransform);
                if (rewriteBoundary == null
                    || !IsUnsafePartialMutationReference(reference!, sr1, sc1, sr2, sc2, cellShiftDirection)) continue;
                throw new InvalidOperationException(
                    $"{operation} cannot preserve partially overlapping workbook source range '{referenceText}'. Edit the complete source range first.");
            }
            foreach (string referenceText in EnumerateMutationPivotSourceReferences(editedSheet)) {
                consumeScannedElement?.Invoke();
                if (!ExcelReference.TryParse(referenceText, out ExcelReference? reference)) continue;
                ValidateMutationReferenceCapacity(reference!, referenceText, operation, capacityTransform);
                if (rewriteBoundary == null
                    || !IsUnsafePartialMutationReference(reference!, sr1, sc1, sr2, sc2, cellShiftDirection)) continue;
                throw new InvalidOperationException(
                    $"{operation} cannot preserve partially overlapping pivot cache source '{referenceText}'. Edit the complete pivot source range first.");
            }
        }

        private static void ValidateMutationReferenceCapacity(
            ExcelReference reference,
            string referenceText,
            string operation,
            Func<ExcelReference, ExcelReference?>? transform) {
            if (transform == null) return;
            try {
                transform(reference);
            } catch (Exception exception) when (exception is ArgumentOutOfRangeException
                || exception is OverflowException) {
                throw new InvalidOperationException(
                    $"{operation} would move reference '{referenceText}' beyond worksheet limits.",
                    exception);
            }
        }

        private static bool IsUnsafePartialMutationReference(
            ExcelReference reference,
            int sr1,
            int sc1,
            int sr2,
            int sc2,
            ExcelCellShiftDirection? cellShiftDirection) {
            reference.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
            bool intersects = rr1 <= sr2 && rr2 >= sr1 && rc1 <= sc2 && rc2 >= sc1;
            bool contained = rr1 >= sr1 && rr2 <= sr2 && rc1 >= sc1 && rc2 <= sc2;
            if (cellShiftDirection == ExcelCellShiftDirection.Left) {
                return intersects && (rr1 < sr1 || rr2 > sr2);
            }
            if (cellShiftDirection == ExcelCellShiftDirection.Up) {
                return intersects && (rc1 < sc1 || rc2 > sc2);
            }
            return intersects && !contained;
        }

        private static IEnumerable<string> EnumerateMutationRangeMetadataReferences(WorksheetPart worksheetPart) {
            IEnumerable<OpenXmlPartRootElement?> roots = new OpenXmlPartRootElement?[] { worksheetPart.Worksheet }
                .Concat(worksheetPart.NamedSheetViewsParts.Select(part => part.NamedSheetViews));
            foreach (OpenXmlPartRootElement? root in roots) {
                if (root == null) continue;
                foreach (OpenXmlElement element in root.Descendants().Prepend(root).Where(IsMutationRangeMetadataElement)) {
                    foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(attribute =>
                        string.Equals(attribute.LocalName, "ref", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(attribute.LocalName, "sqref", StringComparison.OrdinalIgnoreCase))) {
                        foreach (string item in (attribute.Value ?? string.Empty)
                            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                            yield return item;
                        }
                    }
                }
            }
        }

        private IEnumerable<string> EnumerateMutationPivotSourceReferences(ExcelSheet editedSheet) {
            foreach (PivotTableCacheDefinitionPart cachePart in WorkbookPartRoot.PivotTableCacheDefinitionParts) {
                foreach (WorksheetSource source in cachePart.PivotCacheDefinition?.Descendants<WorksheetSource>() ?? Enumerable.Empty<WorksheetSource>()) {
                    if (string.IsNullOrWhiteSpace(source.Id?.Value)
                        && string.Equals(source.Sheet?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase)
                        && source.Reference?.Value is string reference) yield return reference;
                }
                foreach (RangeSet rangeSet in cachePart.PivotCacheDefinition?.Descendants<RangeSet>() ?? Enumerable.Empty<RangeSet>()) {
                    if (string.IsNullOrWhiteSpace(rangeSet.Id?.Value)
                        && string.Equals(rangeSet.Sheet?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase)
                        && rangeSet.Reference?.Value is string reference) yield return reference;
                }
            }
        }

        private IEnumerable<string> EnumerateMutationExternalRangeMetadataReferences(ExcelSheet editedSheet) {
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                foreach (DataReference source in worksheetPart.Worksheet?.Descendants<DataReference>()
                    ?? Enumerable.Empty<DataReference>()) {
                    if (string.IsNullOrWhiteSpace(source.Id?.Value)
                        && string.Equals(source.Sheet?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase)
                        && source.Reference?.Value is string reference) yield return reference;
                }
            }
            foreach (WebPublishItem item in WorkbookRoot.Descendants<WebPublishItem>()) {
                if (item.SourceType?.Value == WebSourceValues.Range
                    && string.Equals(item.SourceObject?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase)
                    && item.SourceRef?.Value is string reference) yield return reference;
            }
        }

        private static bool IsMutationRangeMetadataElement(OpenXmlElement element) =>
            element.GetAttributes().Any(attribute =>
                string.Equals(attribute.LocalName, "sqref", StringComparison.OrdinalIgnoreCase))
            || element is AutoFilter
            || string.Equals(element.LocalName, "nsvFilter", StringComparison.OrdinalIgnoreCase);

        private IEnumerable<MutationFormulaContext> EnumerateMutationFormulaContexts(
            IReadOnlyList<Sheet> sheets,
            int editedSheetIndex,
            string? excludedDefinedName = null) {
            foreach (DefinedName name in WorkbookRoot.DefinedNames?.Elements<DefinedName>()
                ?? Enumerable.Empty<DefinedName>()) {
                if (!string.IsNullOrEmpty(excludedDefinedName)
                    && string.Equals(name.Name?.Value, excludedDefinedName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
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
