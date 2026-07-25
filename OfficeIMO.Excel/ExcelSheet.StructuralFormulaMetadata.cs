using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RewriteWorksheetFormulaReferences(int firstAffectedRow, int rowDelta) {
            RewriteWorkbookCellFormulaReferences(firstAffectedRow, rowDelta, lastDeletedRow: null);
        }

        private void RewriteDeletedWorksheetFormulaReferences(int firstDeletedRow, int lastDeletedRow, int rowDelta) {
            RewriteWorkbookCellFormulaReferences(firstDeletedRow, rowDelta, lastDeletedRow);
        }

        private void RewriteWorkbookCellFormulaReferences(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart
                    || worksheetPart.Worksheet == null) {
                    continue;
                }

                bool isMutatedSheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    || string.Equals(sheetElement.Name?.Value, Name, StringComparison.OrdinalIgnoreCase);
                bool changed = false;
                foreach (Cell cell in worksheetPart.Worksheet.Descendants<Cell>()) {
                    CellFormula? formula = cell.CellFormula;
                    if (formula == null) {
                        continue;
                    }

                    if (formula.Text is string formulaText && formulaText.Length > 0) {
                        string rewritten = lastDeletedRow.HasValue
                            ? RewriteDeletedFormulaReferences(
                                formulaText,
                                firstAffectedRow,
                                lastDeletedRow.Value,
                                rowDelta,
                                Name,
                                rewriteUnqualifiedReferences: isMutatedSheet)
                            : RewriteShiftedFormulaReferences(
                                formulaText,
                                firstAffectedRow,
                                rowDelta,
                                Name,
                                rewriteUnqualifiedReferences: isMutatedSheet);
                        if (!string.Equals(formulaText, rewritten, StringComparison.Ordinal)) {
                            formula.Text = rewritten;
                            changed = true;
                        }
                    }

                    if (isMutatedSheet
                        && formula.Reference?.Value is string formulaReference
                        && TryRemapShiftedReferenceListRows(
                            formulaReference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remappedReferences)) {
                        string? remappedReference = remappedReferences.Count == 0
                            ? null
                            : string.Join(" ", remappedReferences);
                        if (!string.Equals(formulaReference, remappedReference, StringComparison.OrdinalIgnoreCase)) {
                            formula.Reference = remappedReference;
                            changed = true;
                        }
                    }

                    if (isMutatedSheet && formula.FormulaType?.Value == CellFormulaValues.DataTable) {
                        changed |= RemapDataTableInputReference(
                            formula,
                            firstInput: true,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow);
                        changed |= RemapDataTableInputReference(
                            formula,
                            firstInput: false,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow);
                    }
                }

                foreach (DataValidation validation in worksheetPart.Worksheet.Descendants<DataValidation>()) {
                    if (!isMutatedSheet) {
                        changed |= RewriteStructuralFormulaText(
                            validation.Formula1,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
                        changed |= RewriteStructuralFormulaText(
                            validation.Formula2,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
                    }
                }

                if (!isMutatedSheet) {
                    foreach (X14.DataValidation validation in worksheetPart.Worksheet.Descendants<X14.DataValidation>()) {
                        changed |= RewriteStructuralFormulaText(
                            validation.DataValidationForumla1?.Formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
                        changed |= RewriteStructuralFormulaText(
                            validation.DataValidationForumla2?.Formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
                    }
                }

                if (!isMutatedSheet) {
                    foreach (ConditionalFormattingRule rule in worksheetPart.Worksheet.Descendants<ConditionalFormattingRule>()) {
                        foreach (Formula formula in rule.Elements<Formula>()) {
                            changed |= RewriteStructuralFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                rewriteUnqualifiedReferences: false);
                        }
                    }
                }

                foreach (Hyperlink hyperlink in worksheetPart.Worksheet.Descendants<Hyperlink>()) {
                    if (!string.IsNullOrWhiteSpace(hyperlink.Id?.Value)
                        || string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) {
                        continue;
                    }

                    string location = hyperlink.Location!.Value!;
                    string rewrittenLocation = lastDeletedRow.HasValue
                        ? RewriteDeletedFormulaReferences(
                            location,
                            firstAffectedRow,
                            lastDeletedRow.Value,
                            rowDelta,
                            Name,
                            rewriteUnqualifiedReferences: isMutatedSheet)
                        : RewriteShiftedFormulaReferences(
                            location,
                            firstAffectedRow,
                            rowDelta,
                            Name,
                            rewriteUnqualifiedReferences: isMutatedSheet);
                    if (!string.Equals(location, rewrittenLocation, StringComparison.Ordinal)) {
                        hyperlink.Location = rewrittenLocation;
                        changed = true;
                    }
                }

                if (!isMutatedSheet) {
                    foreach (DocumentFormat.OpenXml.Office2010.Excel.Sparkline sparkline
                        in worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>()) {
                        changed |= RewriteStructuralFormulaText(
                            sparkline.Formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
                    }
                }

                if (changed) {
                    worksheetPart.Worksheet.Save();
                }

                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    if (tablePart.Table == null) {
                        continue;
                    }

                    bool tableChanged = false;
                    foreach (CalculatedColumnFormula formula in tablePart.Table.Descendants<CalculatedColumnFormula>()) {
                        tableChanged |= RewriteStructuralFormulaText(
                            formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            isMutatedSheet);
                    }
                    foreach (TotalsRowFormula formula in tablePart.Table.Descendants<TotalsRowFormula>()) {
                        tableChanged |= RewriteStructuralFormulaText(
                            formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            isMutatedSheet);
                    }

                    if (tableChanged) {
                        tablePart.Table.Save();
                    }
                }

                if (!isMutatedSheet && worksheetPart.DrawingsPart != null) {
                    foreach (ChartPart chartPart in worksheetPart.DrawingsPart.ChartParts) {
                        if (chartPart.ChartSpace == null) {
                            continue;
                        }

                        bool chartChanged = false;
                        foreach (DocumentFormat.OpenXml.Drawing.Charts.Formula formula
                            in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                            bool formulaChanged = RewriteStructuralFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                rewriteUnqualifiedReferences: false);
                            chartChanged |= formulaChanged;
                            if (formulaChanged) {
                                InvalidateChartFormulaCache(formula);
                            }
                        }

                        if (chartChanged) {
                            chartPart.ChartSpace.Save();
                        }
                    }
                }
            }

            RewriteChartsheetAndExtendedChartReferences(firstAffectedRow, rowDelta, lastDeletedRow);
        }

        private static bool RemapDataTableInputReference(
            CellFormula formula,
            bool firstInput,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            StringValue? input = firstInput ? formula.R1 : formula.R2;
            if (input?.Value is not string inputReference
                || !TryRemapShiftedReferenceListRows(
                    inputReference,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remapped)) {
                return false;
            }

            if (remapped.Count == 0) {
                if (firstInput) {
                    formula.R1 = null;
                    formula.Input1Deleted = true;
                } else {
                    formula.R2 = null;
                    formula.Input2Deleted = true;
                }
            } else if (firstInput) {
                formula.R1 = remapped[0];
                formula.Input1Deleted = null;
            } else {
                formula.R2 = remapped[0];
                formula.Input2Deleted = null;
            }

            return true;
        }

        private void RewriteChartsheetAndExtendedChartReferences(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                DrawingsPart? drawingsPart = chartsheetPart.DrawingsPart;
                if (drawingsPart == null) {
                    continue;
                }
                foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                    RewriteChartRootReferences(chartPart.ChartSpace, firstAffectedRow, rowDelta, lastDeletedRow);
                }
                foreach (ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                    RewriteChartRootReferences(chartPart.ChartSpace, firstAffectedRow, rowDelta, lastDeletedRow);
                }
            }

            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                if (worksheetPart.DrawingsPart == null) {
                    continue;
                }
                foreach (ExtendedChartPart chartPart in worksheetPart.DrawingsPart.ExtendedChartParts) {
                    RewriteChartRootReferences(chartPart.ChartSpace, firstAffectedRow, rowDelta, lastDeletedRow);
                }
            }
        }

        private void RewriteChartRootReferences(
            OpenXmlPartRootElement? chartRoot,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            if (chartRoot == null) {
                return;
            }

            bool changed = false;
            foreach (OpenXmlLeafTextElement formula in chartRoot.Descendants<OpenXmlLeafTextElement>()
                .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                bool formulaChanged = RewriteStructuralFormulaText(
                    formula,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    rewriteUnqualifiedReferences: false);
                changed |= formulaChanged;
                if (formulaChanged) {
                    InvalidateChartFormulaCache(formula);
                }
            }
            if (changed) {
                chartRoot.Save();
            }
        }

        private static void InvalidateChartFormulaCache(OpenXmlLeafTextElement formula) {
            OpenXmlElement? reference = formula.Parent;
            if (reference == null) {
                return;
            }
            foreach (OpenXmlElement cache in reference.ChildElements
                .Where(element => element.LocalName.EndsWith("Cache", StringComparison.OrdinalIgnoreCase))
                .ToList()) {
                cache.Remove();
            }
        }

        private bool RewriteStructuralFormulaText(
            OpenXmlLeafTextElement? formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            bool rewriteUnqualifiedReferences) {
            if (formula?.Text is not string text || text.Length == 0) {
                return false;
            }

            string rewritten = lastDeletedRow.HasValue
                ? RewriteDeletedFormulaReferences(
                    text,
                    firstAffectedRow,
                    lastDeletedRow.Value,
                    rowDelta,
                    Name,
                    rewriteUnqualifiedReferences)
                : RewriteShiftedFormulaReferences(
                    text,
                    firstAffectedRow,
                    rowDelta,
                    Name,
                    rewriteUnqualifiedReferences);
            if (string.Equals(text, rewritten, StringComparison.Ordinal)) {
                return false;
            }

            formula.Text = rewritten;
            return true;
        }
    }
}
