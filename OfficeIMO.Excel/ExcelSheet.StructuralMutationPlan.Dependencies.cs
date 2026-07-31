using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;
using Xnsv = DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void ClassifyExternalFormulaPlanImpact(
            OpenXmlElement formula,
            ISet<OpenXmlElement> validationImpacts,
            ISet<OpenXmlElement> conditionalFormattingImpacts,
            ISet<OpenXmlElement> sparklineImpacts) {
            for (OpenXmlElement? ancestor = formula.Parent;
                ancestor != null;
                ancestor = ancestor.Parent) {
                if (ancestor is DataValidation || ancestor is X14.DataValidation) {
                    validationImpacts.Add(ancestor);
                    return;
                }
                if (ancestor is ConditionalFormatting || ancestor is X14.ConditionalFormatting) {
                    conditionalFormattingImpacts.Add(ancestor);
                    return;
                }
                if (ancestor is X14.Sparkline || ancestor is X14.SparklineGroup) {
                    sparklineImpacts.Add(ancestor);
                    return;
                }
            }
        }

        private static int CountAffectedRowRecords(
            IReadOnlyList<OpenXmlElement> worksheetElements,
            int firstRow) {
            int count = 0;
            uint previous = 0U;
            foreach (Row row in worksheetElements.OfType<Row>()) {
                uint effective = GetEffectiveRowIndex(row, previous);
                if (effective >= (uint)firstRow
                    || row.RowIndex?.Value is not uint explicitIndex
                    || explicitIndex == 0U) {
                    count++;
                }
                previous = effective;
            }
            return count;
        }

        private int CountWorksheetRangeMetadataPlanImpacts(
            IReadOnlyList<OpenXmlElement> worksheetElements,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            var affected = new HashSet<OpenXmlElement>();
            bool Changes(string? reference) => ReferenceListChangesForPlan(
                reference,
                kind,
                firstRow,
                lastRow,
                count);
            ILookup<OpenXmlElement?, OpenXmlElement> directChildren =
                worksheetElements
                    .Where(element => element is InputCells
                        || element is SortCondition
                        || element is X14.SortCondition)
                    .ToLookup(element => element.Parent);

            foreach (AutoFilter filter in worksheetElements.OfType<AutoFilter>()) {
                if (Changes(filter.Reference?.Value)) {
                    affected.Add(filter);
                }
            }

            SheetDimension? dimension = worksheetElements
                .OfType<SheetDimension>()
                .FirstOrDefault(candidate => candidate.Parent is Worksheet);
            if (dimension != null && Changes(dimension.Reference?.Value)) {
                affected.Add(dimension);
            }

            foreach (IgnoredError error in worksheetElements.OfType<IgnoredError>()) {
                if (Changes(error.SequenceOfReferences?.InnerText)) {
                    affected.Add(error);
                }
            }
            foreach (X14.IgnoredError error in worksheetElements.OfType<X14.IgnoredError>()) {
                if (Changes(error.ReferenceSequence?.Text)) {
                    affected.Add(error);
                }
            }

            Scenarios? scenarios = worksheetElements
                .OfType<Scenarios>()
                .FirstOrDefault(candidate => candidate.Parent is Worksheet);
            if (scenarios != null) {
                if (Changes(scenarios.SequenceOfReferences?.InnerText)) {
                    affected.Add(scenarios);
                }
                foreach (Scenario scenario in worksheetElements
                    .OfType<Scenario>()
                    .Where(candidate => ReferenceEquals(
                        candidate.Parent,
                        scenarios))) {
                    if (directChildren[scenario]
                        .OfType<InputCells>()
                        .Any(input => Changes(input.CellReference?.Value))) {
                        affected.Add(scenario);
                    }
                }
            }

            foreach (CellWatch watch in worksheetElements.OfType<CellWatch>()) {
                if (Changes(watch.CellReference?.Value)) {
                    affected.Add(watch);
                }
            }

            foreach (OpenXmlElement tag in worksheetElements
                .Where(element => string.Equals(
                    element.LocalName,
                    "cellSmartTag",
                    StringComparison.OrdinalIgnoreCase))) {
                string? reference = tag.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(
                        attribute.LocalName,
                        "r",
                        StringComparison.OrdinalIgnoreCase))
                    .Value;
                if (Changes(reference)) {
                    affected.Add(tag);
                }
            }

            foreach (SortState sortState in worksheetElements.OfType<SortState>()) {
                if (Changes(sortState.Reference?.Value)
                    || directChildren[sortState]
                        .OfType<SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))
                    || directChildren[sortState]
                        .OfType<X14.SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))) {
                    affected.Add(sortState);
                }
            }

            foreach (SheetView view in worksheetElements.OfType<SheetView>()) {
                if (Changes(view.TopLeftCell?.Value)) {
                    affected.Add(view);
                }
            }
            foreach (CustomSheetView view in worksheetElements.OfType<CustomSheetView>()) {
                if (Changes(view.TopLeftCell?.Value)) {
                    affected.Add(view);
                }
            }
            foreach (Pane pane in worksheetElements.OfType<Pane>()) {
                if (Changes(pane.TopLeftCell?.Value)) {
                    affected.Add(pane);
                }
            }
            foreach (Selection selection in worksheetElements.OfType<Selection>()) {
                if (Changes(selection.ActiveCell?.Value)
                    || Changes(selection.SequenceOfReferences?.InnerText)) {
                    affected.Add(selection);
                }
            }

            foreach (Break pageBreak in worksheetElements.OfType<Break>()
                .Where(candidate => candidate.Parent is RowBreaks)) {
                if (pageBreak.Id?.Value is uint rowId
                    && rowId > 0U
                    && rowId >= (uint)firstRow) {
                    affected.Add(pageBreak);
                }
            }

            return affected.Count;
        }

        private int CountQueryTableSortPlanImpacts(
            IReadOnlyList<OpenXmlElement> queryElements,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            int impacts = 0;
            ILookup<OpenXmlElement?, OpenXmlElement> directChildren =
                queryElements
                    .Where(element => element is SortCondition
                        || element is X14.SortCondition)
                    .ToLookup(element => element.Parent);
            foreach (SortState sortState in queryElements.OfType<SortState>()) {
                if (ReferenceListChangesForPlan(
                        sortState.Reference?.Value,
                        kind,
                        firstRow,
                        lastRow,
                        count)
                    || directChildren[sortState].OfType<SortCondition>().Any(condition =>
                        ReferenceListChangesForPlan(
                            condition.Reference?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count))
                    || directChildren[sortState].OfType<X14.SortCondition>().Any(condition =>
                        ReferenceListChangesForPlan(
                            condition.Reference?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count))) {
                    impacts++;
                }
            }
            return impacts;
        }

        private bool TableMetadataChangesForPlan(
            Table? table,
            IReadOnlyList<OpenXmlElement> tableElements,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            if (table == null) {
                return false;
            }

            bool Changes(string? reference) => ReferenceListChangesForPlan(
                reference,
                kind,
                firstRow,
                lastRow,
                count);
            ILookup<OpenXmlElement?, OpenXmlElement> directChildren =
                tableElements
                    .Where(element => element is SortCondition
                        || element is X14.SortCondition
                        || element is AutoFilter)
                    .ToLookup(element => element.Parent);

            return Changes(table.Reference?.Value)
                || Changes(directChildren[table]
                    .OfType<AutoFilter>()
                    .FirstOrDefault()
                    ?.Reference?.Value)
                || tableElements.OfType<SortState>().Any(sortState =>
                    Changes(sortState.Reference?.Value)
                    || directChildren[sortState].OfType<SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))
                    || directChildren[sortState].OfType<X14.SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value)));
        }

        private static bool ChartFormulaCacheWillBeInvalidated(
            OpenXmlLeafTextElement formula,
            ILookup<OpenXmlElement?, OpenXmlElement> directChildren) {
            OpenXmlElement? reference = formula.Parent;
            return reference != null
                && (directChildren[reference].Any(element =>
                        element.LocalName.EndsWith("Cache", StringComparison.OrdinalIgnoreCase))
                    || (string.Equals(reference.LocalName, "numDim", StringComparison.Ordinal)
                        || string.Equals(reference.LocalName, "strDim", StringComparison.Ordinal))
                    && directChildren[reference].Any(element =>
                        string.Equals(element.LocalName, "lvl", StringComparison.Ordinal)));
        }

        private int CountNamedSheetViewPlanImpacts(
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            MutationPlanScanBudget budget) {
            int impacts = 0;
            foreach (NamedSheetViewsPart part in _worksheetPart.NamedSheetViewsParts) {
                budget.Consume();
                Xnsv.NamedSheetViews? views = part.NamedSheetViews;
                if (views == null) {
                    continue;
                }

                budget.Consume();
                foreach (OpenXmlElement element in views.Descendants()) {
                    budget.Consume();
                    if (element is Xnsv.NsvFilter filter
                        && ReferenceListChangesForPlan(
                            filter.Ref?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        impacts++;
                    }
                }
            }
            return impacts;
        }

        private int CountTableFormulaPlanImpacts(
            TableDefinitionPart tablePart,
            bool rewriteUnqualifiedReferences,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            MutationPlanScanBudget budget,
            out IReadOnlyList<OpenXmlElement> tableElements) {
            budget.Consume();
            Table? table = tablePart.Table;
            if (table == null) {
                tableElements = Array.Empty<OpenXmlElement>();
                return 0;
            }

            int rowDelta = kind == ExcelRowMutationKind.Insert ? count : -count;
            int? lastDeletedRow = kind == ExcelRowMutationKind.Delete ? lastRow : null;
            GetTableFormulaAnchorDeltas(
                table,
                firstRow,
                rowDelta,
                lastDeletedRow,
                out int calculatedColumnAnchorDelta,
                out int totalsRowAnchorDelta,
                out int calculatedColumnAnchorRow,
                out int totalsRowAnchorRow);

            int impacts = 0;
            var inspectedElements = new List<OpenXmlElement>();
            foreach (OpenXmlElement element in table.Descendants()) {
                budget.Consume();
                inspectedElements.Add(element);
                bool changes;
                if (element is CalculatedColumnFormula calculatedColumnFormula) {
                    changes = rewriteUnqualifiedReferences
                        ? AnchoredFormulaChangesForPlan(
                            calculatedColumnFormula,
                            firstRow,
                            rowDelta,
                            lastDeletedRow,
                            calculatedColumnAnchorDelta,
                            calculatedColumnAnchorRow)
                        : FormulaChangesForPlan(
                            calculatedColumnFormula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences: false);
                } else if (element is TotalsRowFormula totalsRowFormula) {
                    changes = rewriteUnqualifiedReferences
                        ? AnchoredFormulaChangesForPlan(
                            totalsRowFormula,
                            firstRow,
                            rowDelta,
                            lastDeletedRow,
                            totalsRowAnchorDelta,
                            totalsRowAnchorRow)
                        : FormulaChangesForPlan(
                            totalsRowFormula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences: false);
                } else {
                    changes = element is OpenXmlLeafTextElement formula
                        && IsStructuralFormulaElement(formula)
                        && FormulaChangesForPlan(
                            formula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences);
                }

                if (changes) {
                    impacts++;
                }
            }
            tableElements = inspectedElements;
            return impacts;
        }

        private int CountAnchoredMetadataFormulaPlanImpacts(
            OpenXmlElement metadata,
            IReadOnlyDictionary<OpenXmlElement, List<string?>>
                anchoredMetadataFormulaTexts,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            string? references;
            switch (metadata) {
                case DataValidation validation:
                    references = validation.SequenceOfReferences?.InnerText;
                    break;
                case X14.DataValidation validation:
                    references = validation.ReferenceSequence?.Text;
                    break;
                case ConditionalFormatting formatting:
                    references = formatting.SequenceOfReferences?.InnerText;
                    break;
                case X14.ConditionalFormatting formatting:
                    references = formatting.GetFirstChild<Xm.ReferenceSequence>()?.Text;
                    break;
                default:
                    return 0;
            }
            IReadOnlyList<string?> formulaTexts =
                anchoredMetadataFormulaTexts.TryGetValue(
                    metadata,
                    out List<string?>? recordedFormulaTexts)
                    ? recordedFormulaTexts
                    : Array.Empty<string?>();

            if (string.IsNullOrWhiteSpace(references)
                || !TryGetReferenceListAnchorRow(references!, out int oldAnchorRow)) {
                return 0;
            }

            int rowDelta = kind == ExcelRowMutationKind.Insert ? count : -count;
            int? lastDeletedRow = kind == ExcelRowMutationKind.Delete ? lastRow : null;
            string updatedReferences = references!;
            if (TryRemapShiftedReferenceListRows(
                    references!,
                    firstRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remapped)) {
                if (remapped.Count == 0) {
                    return formulaTexts.Count(formulaText =>
                        !string.IsNullOrEmpty(formulaText));
                }
                updatedReferences = string.Join(" ", remapped);
            }

            if (!TryGetReferenceListAnchorRow(updatedReferences, out int newAnchorRow)) {
                return 0;
            }

            int anchorRowDelta = newAnchorRow - oldAnchorRow;
            int relativeFormulaSourceRowDelta = GetRelativeFormulaSourceRowDelta(
                oldAnchorRow,
                newAnchorRow,
                firstRow,
                rowDelta,
                lastDeletedRow);
            return formulaTexts.Count(formulaText =>
                AnchoredFormulaChangesForPlan(
                    formulaText,
                    firstRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    relativeReferencesFollowAnchor: false,
                    relativeFormulaSourceRowDelta,
                    relativeFormulaAnchorRow: null));
        }

        private static void AddAnchoredMetadataFormulaText(
            IDictionary<OpenXmlElement, List<string?>> formulaTextsByMetadata,
            OpenXmlElement element) {
            string? formulaText;
            switch (element) {
                case Formula1 formula1:
                    formulaText = formula1.Text;
                    break;
                case Formula2 formula2:
                    formulaText = formula2.Text;
                    break;
                case Formula formula:
                    formulaText = formula.Text;
                    break;
                case Xm.Formula xmFormula:
                    formulaText = xmFormula.Text;
                    break;
                case ConditionalFormatValueObject threshold
                    when threshold.Type?.Value ==
                        ConditionalFormatValueObjectValues.Formula:
                    formulaText = threshold.Val?.Value;
                    break;
                default:
                    return;
            }

            OpenXmlElement? metadata = element.Parent;
            while (metadata != null
                && metadata is not DataValidation
                && metadata is not X14.DataValidation
                && metadata is not ConditionalFormatting
                && metadata is not X14.ConditionalFormatting) {
                metadata = metadata.Parent;
            }
            if (metadata == null) {
                return;
            }

            if (!formulaTextsByMetadata.TryGetValue(
                    metadata,
                    out List<string?>? formulaTexts)) {
                formulaTexts = new List<string?>();
                formulaTextsByMetadata.Add(metadata, formulaTexts);
            }
            formulaTexts.Add(formulaText);
        }

        private int CountCommentPlanImpacts(
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            MutationPlanScanBudget budget) {
            int comments = 0;
            var impactedLegacyCommentCells = new HashSet<(int Row, int Column)>();
            WorksheetCommentsPart? legacyPart = _worksheetPart.WorksheetCommentsPart;
            if (legacyPart?.Comments != null) {
                budget.Consume();
                budget.Consume();
                foreach (OpenXmlElement element in legacyPart.Comments.Descendants()) {
                    budget.Consume();
                    if (element is Comment comment
                        && ReferenceListChangesForPlan(
                            comment.Reference?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        comments++;
                        if (comment.Reference?.Value is string reference
                            && TryParseReference(reference, out var bounds)
                            && bounds.r1 == bounds.r2
                            && bounds.c1 == bounds.c2) {
                            impactedLegacyCommentCells.Add((bounds.r1, bounds.c1));
                        }
                    }
                }
            }

            VmlDrawingPart? commentVmlPart = TryGetCommentVmlPart();
            if (commentVmlPart != null) {
                budget.Consume();
                ConsumeVmlElementsForMutationPlan(commentVmlPart, budget);
                XDocument document = LoadOrCreateVmlDocument(commentVmlPart);
                XNamespace vmlNamespace = "urn:schemas-microsoft-com:vml";
                XNamespace excelNamespace =
                    "urn:schemas-microsoft-com:office:excel";
                XElement? root = document.Root;
                var shapes = new List<XElement>();
                if (root != null) {
                    foreach (XElement element in root.Descendants()) {
                        if (element.Name == vmlNamespace + "shape"
                            && ReferenceEquals(element.Parent, root)) {
                            shapes.Add(element);
                        }
                    }
                }
                foreach (XElement shape in shapes) {
                    XElement? clientData =
                        shape.Element(excelNamespace + "ClientData");
                    if (clientData == null
                        || !TryParseVmlCoordinate(
                            clientData.Element(excelNamespace + "Row")?.Value,
                            out int zeroBasedRow)
                        || !TryParseVmlCoordinate(
                            clientData.Element(excelNamespace + "Column")?.Value,
                            out int zeroBasedColumn)) {
                        continue;
                    }

                    XElement? anchor =
                        clientData.Element(excelNamespace + "Anchor");
                    if (anchor == null
                        || !RemapVmlAnchorRows(
                            new XElement(anchor),
                            firstRow,
                            kind == ExcelRowMutationKind.Insert ? count : -count,
                            kind == ExcelRowMutationKind.Delete ? lastRow : (int?)null,
                            columnDelta: 0,
                            GetVmlAnchorPlacement(
                                clientData,
                                excelNamespace))) {
                        continue;
                    }

                    if (impactedLegacyCommentCells.Add(
                        (zeroBasedRow + 1, zeroBasedColumn + 1))) {
                        comments++;
                    }
                }
            }

            foreach (WorksheetThreadedCommentsPart threadedPart in
                _worksheetPart.WorksheetThreadedCommentsParts) {
                budget.Consume();
                Threaded.ThreadedComments? root = threadedPart.ThreadedComments;
                if (root == null) {
                    continue;
                }

                budget.Consume();
                List<Threaded.ThreadedComment> allComments = root
                    .Elements<Threaded.ThreadedComment>()
                    .ToList();
                var removedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var childrenByParentId =
                    new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (Threaded.ThreadedComment comment in allComments) {
                    budget.Consume();
                    if (comment.Id?.Value is string commentId
                        && comment.ParentId?.Value is string parentId) {
                        if (!childrenByParentId.TryGetValue(
                                parentId,
                                out List<string>? children)) {
                            children = new List<string>();
                            childrenByParentId[parentId] = children;
                        }
                        children.Add(commentId);
                    }
                    if (comment.Ref?.Value is string reference
                        && kind == ExcelRowMutationKind.Delete
                        && TryParseReference(reference, out var bounds)
                        && TryRemapShiftedReferenceRows(
                            bounds,
                            firstRow,
                            -count,
                            lastRow,
                            out var remapped)
                        && remapped == null
                        && comment.Id?.Value is string id) {
                        removedIds.Add(id);
                    }
                }
                var pendingRemovedParents = new Queue<string>(removedIds);
                while (pendingRemovedParents.Count > 0) {
                    string removedParent = pendingRemovedParents.Dequeue();
                    if (!childrenByParentId.TryGetValue(
                            removedParent,
                            out List<string>? children)) {
                        continue;
                    }
                    foreach (string childId in children) {
                        if (removedIds.Add(childId)) {
                            pendingRemovedParents.Enqueue(childId);
                        }
                    }
                }

                foreach (Threaded.ThreadedComment comment in allComments) {
                    if ((comment.Id?.Value is string id && removedIds.Contains(id))
                        || ReferenceListChangesForPlan(
                            comment.Ref?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        comments++;
                    }
                }
            }
            return comments;
        }

        private static bool UsesAnchoredTargetFormulaSemantics(OpenXmlElement formula) {
            for (OpenXmlElement? parent = formula.Parent; parent != null; parent = parent.Parent) {
                if (parent is DataValidation
                    || parent is X14.DataValidation
                    || parent is ConditionalFormatting
                    || parent is X14.ConditionalFormatting) {
                    return true;
                }
            }
            return false;
        }

        private bool AnchoredFormulaChangesForPlan(
            OpenXmlLeafTextElement formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            int relativeFormulaAnchorRow) {
            return AnchoredFormulaChangesForPlan(
                formula.Text,
                firstAffectedRow,
                rowDelta,
                lastDeletedRow,
                anchorRowDelta,
                relativeReferencesFollowAnchor: true,
                relativeFormulaSourceRowDelta: 0,
                relativeFormulaAnchorRow);
        }

        private bool AnchoredFormulaChangesForPlan(
            string? text,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            bool relativeReferencesFollowAnchor,
            int relativeFormulaSourceRowDelta,
            int? relativeFormulaAnchorRow) {
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            string rewritten = RewriteAnchoredFormulaReferences(
                text!,
                firstAffectedRow,
                rowDelta,
                lastDeletedRow,
                Name,
                anchorRowDelta,
                relativeReferencesFollowAnchor,
                relativeFormulaSourceRowDelta,
                relativeFormulaAnchorRow);
            return !string.Equals(text, rewritten, StringComparison.Ordinal);
        }
    }
}
