using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

        private bool AnchoredFormulaChangesForPlan(
            OpenXmlLeafTextElement formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            int relativeFormulaAnchorRow) {
            string text = formula.Text;
            string rewritten = RewriteAnchoredFormulaReferences(
                text,
                firstAffectedRow,
                rowDelta,
                lastDeletedRow,
                Name,
                anchorRowDelta,
                relativeReferencesFollowAnchor: true,
                relativeFormulaSourceRowDelta: 0,
                relativeFormulaAnchorRow);
            return !string.Equals(text, rewritten, StringComparison.Ordinal);
        }
    }
}
