using System;
using System.Collections.Generic;
using System.Linq;
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
                if (ancestor is X14.Sparkline) {
                    sparklineImpacts.Add(ancestor);
                    return;
                }
            }
        }

        private static int CountAffectedRowRecords(Worksheet worksheet, int firstRow) {
            int count = 0;
            uint previous = 0U;
            foreach (Row row in worksheet.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>()) {
                uint effective = GetEffectiveRowIndex(row, previous);
                if (effective >= (uint)firstRow) {
                    count++;
                }
                previous = effective;
            }
            return count;
        }

        private int CountWorksheetRangeMetadataPlanImpacts(
            Worksheet worksheet,
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

            foreach (AutoFilter filter in worksheet.Descendants<AutoFilter>()) {
                if (Changes(filter.Reference?.Value)) {
                    affected.Add(filter);
                }
            }

            SheetDimension? dimension = worksheet.GetFirstChild<SheetDimension>();
            if (dimension != null && Changes(dimension.Reference?.Value)) {
                affected.Add(dimension);
            }

            foreach (IgnoredError error in worksheet.Descendants<IgnoredError>()) {
                if (Changes(error.SequenceOfReferences?.InnerText)) {
                    affected.Add(error);
                }
            }
            foreach (X14.IgnoredError error in worksheet.Descendants<X14.IgnoredError>()) {
                if (Changes(error.ReferenceSequence?.Text)) {
                    affected.Add(error);
                }
            }

            Scenarios? scenarios = worksheet.GetFirstChild<Scenarios>();
            if (scenarios != null) {
                if (Changes(scenarios.SequenceOfReferences?.InnerText)) {
                    affected.Add(scenarios);
                }
                foreach (Scenario scenario in scenarios.Elements<Scenario>()) {
                    if (scenario.Elements<InputCells>()
                        .Any(input => Changes(input.CellReference?.Value))) {
                        affected.Add(scenario);
                    }
                }
            }

            foreach (CellWatch watch in worksheet.Descendants<CellWatch>()) {
                if (Changes(watch.CellReference?.Value)) {
                    affected.Add(watch);
                }
            }

            foreach (OpenXmlElement tag in worksheet.Descendants()
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

            foreach (SortState sortState in worksheet.Descendants<SortState>()) {
                if (Changes(sortState.Reference?.Value)
                    || sortState.Elements<SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))
                    || sortState.Elements<X14.SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))) {
                    affected.Add(sortState);
                }
            }

            foreach (SheetView view in worksheet.Descendants<SheetView>()) {
                if (Changes(view.TopLeftCell?.Value)) {
                    affected.Add(view);
                }
            }
            foreach (CustomSheetView view in worksheet.Descendants<CustomSheetView>()) {
                if (Changes(view.TopLeftCell?.Value)) {
                    affected.Add(view);
                }
            }
            foreach (Pane pane in worksheet.Descendants<Pane>()) {
                if (Changes(pane.TopLeftCell?.Value)) {
                    affected.Add(pane);
                }
            }
            foreach (Selection selection in worksheet.Descendants<Selection>()) {
                if (Changes(selection.ActiveCell?.Value)
                    || Changes(selection.SequenceOfReferences?.InnerText)) {
                    affected.Add(selection);
                }
            }

            foreach (Break pageBreak in worksheet.Descendants<RowBreaks>()
                .SelectMany(rowBreaks => rowBreaks.Elements<Break>())) {
                if (pageBreak.Id?.Value is uint rowId
                    && rowId > 0U
                    && rowId >= (uint)firstRow) {
                    affected.Add(pageBreak);
                }
            }

            return affected.Count;
        }

        private int CountQueryTableSortPlanImpacts(
            QueryTable? queryTable,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            if (queryTable == null) {
                return 0;
            }

            int impacts = 0;
            foreach (SortState sortState in queryTable.Descendants<SortState>()) {
                if (ReferenceListChangesForPlan(
                        sortState.Reference?.Value,
                        kind,
                        firstRow,
                        lastRow,
                        count)
                    || sortState.Elements<SortCondition>().Any(condition =>
                        ReferenceListChangesForPlan(
                            condition.Reference?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count))
                    || sortState.Elements<X14.SortCondition>().Any(condition =>
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

            return Changes(table.Reference?.Value)
                || Changes(table.GetFirstChild<AutoFilter>()?.Reference?.Value)
                || table.Descendants<SortState>().Any(sortState =>
                    Changes(sortState.Reference?.Value)
                    || sortState.Elements<SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value))
                    || sortState.Elements<X14.SortCondition>()
                        .Any(condition => Changes(condition.Reference?.Value)));
        }

        private static bool ChartFormulaCacheWillBeInvalidated(OpenXmlLeafTextElement formula) {
            OpenXmlElement? reference = formula.Parent;
            return reference != null
                && (reference.ChildElements.Any(element =>
                        element.LocalName.EndsWith("Cache", StringComparison.OrdinalIgnoreCase))
                    || (string.Equals(reference.LocalName, "numDim", StringComparison.Ordinal)
                        || string.Equals(reference.LocalName, "strDim", StringComparison.Ordinal))
                    && reference.ChildElements.Any(element =>
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
            MutationPlanScanBudget budget) {
            budget.Consume();
            Table? table = tablePart.Table;
            if (table == null) {
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
            foreach (OpenXmlElement element in table.Descendants()) {
                budget.Consume();
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
            return impacts;
        }

        private int CountAnchoredMetadataFormulaPlanImpacts(
            OpenXmlElement metadata,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            string? references;
            IEnumerable<string?> formulaTexts;
            switch (metadata) {
                case DataValidation validation:
                    references = validation.SequenceOfReferences?.InnerText;
                    formulaTexts = new[] {
                        validation.Formula1?.Text,
                        validation.Formula2?.Text
                    };
                    break;
                case X14.DataValidation validation:
                    references = validation.ReferenceSequence?.Text;
                    formulaTexts = new[] {
                        validation.DataValidationForumla1?.Formula?.Text,
                        validation.DataValidationForumla2?.Formula?.Text
                    };
                    break;
                case ConditionalFormatting formatting:
                    references = formatting.SequenceOfReferences?.InnerText;
                    formulaTexts = formatting.Descendants<Formula>()
                        .Select(formula => formula.Text)
                        .Concat(formatting.Descendants<ConditionalFormatValueObject>()
                            .Where(threshold =>
                                threshold.Type?.Value == ConditionalFormatValueObjectValues.Formula)
                            .Select(threshold => threshold.Val?.Value));
                    break;
                case X14.ConditionalFormatting formatting:
                    references = formatting.GetFirstChild<Xm.ReferenceSequence>()?.Text;
                    formulaTexts = formatting.Descendants<Xm.Formula>()
                        .Select(formula => formula.Text);
                    break;
                default:
                    return 0;
            }

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
                    return 0;
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

        private int CountCommentPlanImpacts(
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            MutationPlanScanBudget budget) {
            int comments = 0;
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
