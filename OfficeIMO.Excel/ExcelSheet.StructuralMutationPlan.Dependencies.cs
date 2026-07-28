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

        private int CountCommentPlanImpacts(MutationPlanScanBudget budget) {
            int comments = 0;
            WorksheetCommentsPart? legacyPart = _worksheetPart.WorksheetCommentsPart;
            if (legacyPart?.Comments != null) {
                budget.Consume();
                budget.Consume();
                foreach (OpenXmlElement element in legacyPart.Comments.Descendants()) {
                    budget.Consume();
                    if (element is Comment) {
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
                foreach (OpenXmlElement element in root.Descendants()) {
                    budget.Consume();
                    if (element is Threaded.ThreadedComment) {
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
