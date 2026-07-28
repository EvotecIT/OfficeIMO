using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Validates and describes an insertion without changing the workbook.
        /// </summary>
        public ExcelRowMutationPlan PlanInsertRows(
            int firstRow,
            int count = 1,
            ExcelMutationPlanOptions? options = null) {
            ValidateStructuralRowArguments(firstRow, count);
            ExcelMutationPlanOptions effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                ExcelRowMutationPlan plan = BuildRowMutationPlan(
                    ExcelRowMutationKind.Insert,
                    firstRow,
                    count,
                    effective);
                PreflightRowInsertion(firstRow, count);
                return plan;
            });
        }

        /// <summary>
        /// Validates and describes a deletion without changing the workbook.
        /// </summary>
        public ExcelRowMutationPlan PlanDeleteRows(
            int firstRow,
            int count = 1,
            ExcelMutationPlanOptions? options = null) {
            ValidateStructuralRowArguments(firstRow, count);
            ExcelMutationPlanOptions effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                ExcelRowMutationPlan plan = BuildRowMutationPlan(
                    ExcelRowMutationKind.Delete,
                    firstRow,
                    count,
                    effective);
                PreflightRowDeletion(firstRow, count);
                return plan;
            });
        }

        private ExcelRowMutationPlan BuildRowMutationPlan(
            ExcelRowMutationKind kind,
            int firstRow,
            int count,
            ExcelMutationPlanOptions effective) {
            var budget = new MutationPlanScanBudget(effective.MaximumScannedElements);
            var impacts = new List<ExcelMutationImpact>();
            int lastRow = firstRow + count - 1;
            int cells = 0;
            int formulas = 0;
            int definedNames = 0;
            int validation = 0;
            int conditionalFormatting = 0;
            int mergedCells = 0;
            int hyperlinks = 0;
            int sparklines = 0;
            var targetCellCoordinates = new HashSet<long>();
            int mutatedSheetIndex = -1;
            int sheetIndex = 0;

            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                budget.Consume();
                if (mutatedSheetIndex < 0
                    && string.Equals(sheetElement.Name?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                    mutatedSheetIndex = sheetIndex;
                }
                sheetIndex++;
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart
                    || worksheetPart.Worksheet == null) {
                    continue;
                }

                bool rewriteUnqualified = ReferenceEquals(worksheetPart, _worksheetPart)
                    || string.Equals(sheetElement.Name?.Value, Name, StringComparison.OrdinalIgnoreCase);
                foreach (OpenXmlElement element in worksheetPart.Worksheet.Descendants()) {
                    budget.Consume();
                    if (element is OpenXmlLeafTextElement formula
                        && IsStructuralFormulaElement(formula)
                        && FormulaChangesForPlan(
                            formula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualified)) {
                        formulas++;
                    }
                    if (!rewriteUnqualified) {
                        continue;
                    }
                    if (element is Cell cell
                        && A1.TryParseCellReferenceFast(cell.CellReference?.Value, out int row, out int column)) {
                        targetCellCoordinates.Add(((long)row << 32) | (uint)column);
                        if (row >= firstRow) {
                            cells++;
                        }
                    } else if (element is DataValidation) {
                        validation++;
                    } else if (element is ConditionalFormatting) {
                        conditionalFormatting++;
                    } else if (element is MergeCell) {
                        mergedCells++;
                    } else if (element is Hyperlink) {
                        hyperlinks++;
                    } else if (element is DocumentFormat.OpenXml.Office2010.Excel.Sparkline) {
                        sparklines++;
                    }
                }

                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    budget.Consume();
                    foreach (OpenXmlElement element in tablePart.Table?.Descendants()
                        ?? Enumerable.Empty<OpenXmlElement>()) {
                        budget.Consume();
                        if (element is OpenXmlLeafTextElement formula
                            && IsStructuralFormulaElement(formula)
                            && FormulaChangesForPlan(
                            formula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualified)) {
                            formulas++;
                        }
                    }
                }
            }

            if (_pendingCellValueDirectSaveBuffer != null) {
                foreach ((int Row, int Column, object? Value) pending in
                    _pendingCellValueDirectSaveBuffer.EnumerateWrittenCells()) {
                    budget.Consume();
                    long coordinate = ((long)pending.Row << 32) | (uint)pending.Column;
                    if (!targetCellCoordinates.Add(coordinate)) {
                        continue;
                    }
                    if (pending.Row >= firstRow) {
                        cells++;
                    }
                    if (pending.Value is DirectFormulaCellValue pendingFormula
                        && FormulaChangesForPlan(
                            pendingFormula.Formula,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences: true)) {
                        formulas++;
                    }
                }
            }

            foreach (OpenXmlElement element in WorkbookRoot.Descendants()) {
                budget.Consume();
                if (element is DefinedName definedName
                    && FormulaChangesForPlan(
                    definedName.Text,
                    kind,
                    firstRow,
                    lastRow,
                    count,
                    rewriteUnqualifiedReferences:
                        mutatedSheetIndex >= 0
                        && definedName.LocalSheetId?.Value == (uint)mutatedSheetIndex)) {
                    definedNames++;
                }
            }
            AddImpact(
                impacts,
                "worksheet-cells",
                cells,
                "Cells at or below the structural boundary can move or be removed.");
            AddImpact(
                impacts,
                "formula-references",
                formulas,
                "Formula-bearing cells and metadata whose references will be rewritten.");
            AddImpact(
                impacts,
                "defined-names",
                definedNames,
                "Workbook or worksheet names whose formulas will be rewritten.");

            AddImpact(
                impacts,
                "tables",
                CountBounded(_worksheetPart.TableDefinitionParts, budget),
                "Worksheet tables are checked for range, filter, and calculated-column changes.");
            AddImpact(
                impacts,
                "validation",
                validation,
                "Data-validation ranges and formulas are checked and remapped.");
            AddImpact(
                impacts,
                "conditional-formatting",
                conditionalFormatting,
                "Conditional-format ranges and formulas are checked and remapped.");
            AddImpact(
                impacts,
                "merged-cells",
                mergedCells,
                "Merged ranges crossing or following the boundary are remapped.");
            AddImpact(
                impacts,
                "hyperlinks",
                hyperlinks,
                "Internal link destinations and cell anchors are checked and remapped.");
            AddImpact(
                impacts,
                "drawings",
                CountBounded(_worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants()
                    ?? Enumerable.Empty<OpenXmlElement>(), budget),
                "Drawing anchors, shapes, and embedded chart locations are checked and remapped.");
            AddImpact(
                impacts,
                "pivots",
                CountBounded(_worksheetPart.PivotTableParts, budget),
                "Pivot output and source boundaries are validated before mutation.");
            AddImpact(
                impacts,
                "comments",
                CountBounded(
                    _worksheetPart.WorksheetCommentsPart?.Comments?.Descendants()
                    ?? Enumerable.Empty<Comment>(),
                    budget),
                "Legacy comment anchors are checked and remapped.");
            AddImpact(
                impacts,
                "sparklines",
                sparklines,
                "Sparkline locations and data references are checked and remapped.");

            return new ExcelRowMutationPlan(
                this,
                kind,
                Name,
                firstRow,
                count,
                budget.Scanned,
                impacts);
        }

        private void EnsureMutationPlanCanInspectWithoutMaterializing() {
            if (_excelDocument.HasDeferredDirectDataSetImport) {
                throw new InvalidOperationException(
                    "A non-mutating structural plan cannot inspect pending deferred worksheet writes. " +
                    "Materialize or save those writes before requesting the plan.");
            }
        }

        private bool FormulaChangesForPlan(
            string? formula,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            bool rewriteUnqualifiedReferences) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return false;
            }

            string rewritten = kind == ExcelRowMutationKind.Insert
                ? RewriteShiftedFormulaReferences(
                    formula!,
                    firstRow,
                    count,
                    Name,
                    rewriteUnqualifiedReferences)
                : RewriteDeletedFormulaReferences(
                    formula!,
                    firstRow,
                    lastRow,
                    -count,
                    Name,
                    rewriteUnqualifiedReferences);
            return !string.Equals(formula, rewritten, StringComparison.Ordinal);
        }

        private static bool IsStructuralFormulaElement(OpenXmlLeafTextElement element) =>
            element is CellFormula
            || element is Formula
            || element is Formula1
            || element is Formula2
            || string.Equals(element.LocalName, "f", StringComparison.Ordinal);

        private static int CountBounded<T>(
            IEnumerable<T> items,
            MutationPlanScanBudget budget) {
            int count = 0;
            foreach (T _ in items) {
                budget.Consume();
                count++;
            }
            return count;
        }

        private static void AddImpact(
            ICollection<ExcelMutationImpact> impacts,
            string category,
            int count,
            string description) {
            if (count > 0) {
                impacts.Add(new ExcelMutationImpact(category, count, description));
            }
        }

        private sealed class MutationPlanScanBudget {
            private readonly int _maximum;

            internal MutationPlanScanBudget(int maximum) {
                _maximum = maximum;
            }

            internal int Scanned { get; private set; }

            internal void Consume() {
                if (Scanned >= _maximum) {
                    throw new InvalidOperationException(
                        $"Excel mutation impact analysis exceeded its limit of {_maximum} inspected elements. " +
                        "Increase ExcelMutationPlanOptions.MaximumScannedElements explicitly for this workbook.");
                }

                Scanned++;
            }
        }
    }
}
