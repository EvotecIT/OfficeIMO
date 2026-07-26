using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xm = DocumentFormat.OpenXml.Office.Excel;

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
                        foreach (ConditionalFormatValueObject threshold in rule
                            .Descendants<ConditionalFormatValueObject>()
                            .Where(item => item.Type?.Value == ConditionalFormatValueObjectValues.Formula)) {
                            if (threshold.Val?.Value is not string formulaText || formulaText.Length == 0) {
                                continue;
                            }

                            string rewritten = lastDeletedRow.HasValue
                                ? RewriteDeletedFormulaReferences(
                                    formulaText,
                                    firstAffectedRow,
                                    lastDeletedRow.Value,
                                    rowDelta,
                                    Name,
                                    rewriteUnqualifiedReferences: false)
                                : RewriteShiftedFormulaReferences(
                                    formulaText,
                                    firstAffectedRow,
                                    rowDelta,
                                    Name,
                                    rewriteUnqualifiedReferences: false);
                            if (!string.Equals(formulaText, rewritten, StringComparison.Ordinal)) {
                                threshold.Val = rewritten;
                                changed = true;
                            }
                        }
                    }

                    foreach (Xm.Formula formula in worksheetPart.Worksheet
                        .Descendants<X14.ConditionalFormatting>()
                        .SelectMany(formatting => formatting.Descendants<Xm.Formula>())) {
                        changed |= RewriteStructuralFormulaText(
                            formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            rewriteUnqualifiedReferences: false);
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
                foreach (X14.SparklineGroup group
                    in worksheetPart.Worksheet.Descendants<X14.SparklineGroup>()) {
                    changed |= RewriteStructuralFormulaText(
                        group.Formula,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        rewriteUnqualifiedReferences: isMutatedSheet);
                }

                if (changed) {
                    worksheetPart.Worksheet.Save();
                }

                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    if (tablePart.Table == null) {
                        continue;
                    }

                    bool tableChanged = false;
                    GetTableFormulaAnchorDeltas(
                        tablePart.Table,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out int calculatedColumnAnchorDelta,
                        out int totalsRowAnchorDelta,
                        out int calculatedColumnAnchorRow,
                        out int totalsRowAnchorRow);
                    foreach (CalculatedColumnFormula formula in tablePart.Table.Descendants<CalculatedColumnFormula>()) {
                        tableChanged |= isMutatedSheet
                            ? RewriteAnchoredFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                calculatedColumnAnchorDelta,
                                relativeReferencesFollowAnchor: true,
                                relativeFormulaAnchorRow: calculatedColumnAnchorRow)
                            : RewriteStructuralFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                rewriteUnqualifiedReferences: false);
                    }
                    foreach (TotalsRowFormula formula in tablePart.Table.Descendants<TotalsRowFormula>()) {
                        tableChanged |= isMutatedSheet
                            ? RewriteAnchoredFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                totalsRowAnchorDelta,
                                relativeReferencesFollowAnchor: true,
                                relativeFormulaAnchorRow: totalsRowAnchorRow)
                            : RewriteStructuralFormulaText(
                                formula,
                                firstAffectedRow,
                                rowDelta,
                                lastDeletedRow,
                                rewriteUnqualifiedReferences: false);
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
                        foreach (OpenXmlLeafTextElement formula in chartPart.ChartSpace
                            .Descendants<OpenXmlLeafTextElement>()
                            .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
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

                RewriteDrawingShapeTextLinks(
                    worksheetPart.DrawingsPart,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    rewriteUnqualifiedReferences: isMutatedSheet);
            }

            RewriteChartsheetAndExtendedChartReferences(firstAffectedRow, rowDelta, lastDeletedRow);
        }

        private static void GetTableFormulaAnchorDeltas(
            Table table,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            out int calculatedColumnAnchorDelta,
            out int totalsRowAnchorDelta,
            out int calculatedColumnAnchorRow,
            out int totalsRowAnchorRow) {
            calculatedColumnAnchorDelta = 0;
            totalsRowAnchorDelta = 0;
            calculatedColumnAnchorRow = 0;
            totalsRowAnchorRow = 0;
            if (table.Reference?.Value is not string reference
                || !A1.TryParseRange(
                    reference.Replace("$", string.Empty),
                    out int oldFirstRow,
                    out _,
                    out int oldLastRow,
                    out _)) {
                return;
            }

            bool hasHeaderRow = (table.HeaderRowCount?.Value ?? 1U) > 0U;
            calculatedColumnAnchorRow = oldFirstRow + (hasHeaderRow ? 1 : 0);
            totalsRowAnchorRow = oldLastRow;

            int newFirstRow = oldFirstRow;
            int newLastRow = oldLastRow;
            if (TryRemapShiftedReferenceListRows(
                    reference,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remapped)
                && remapped.Count > 0
                && A1.TryParseRange(
                    remapped[0].Replace("$", string.Empty),
                    out int remappedFirstRow,
                    out _,
                    out int remappedLastRow,
                    out _)) {
                newFirstRow = remappedFirstRow;
                newLastRow = remappedLastRow;
            }

            calculatedColumnAnchorDelta = newFirstRow - oldFirstRow;
            totalsRowAnchorDelta = newLastRow - oldLastRow;
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
                RewriteDrawingShapeTextLinks(
                    drawingsPart,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    rewriteUnqualifiedReferences: false);
                foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                    RewriteChartRootReferences(
                        chartPart.ChartSpace,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        rewriteUnqualifiedReferences: false);
                }
                foreach (ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                    RewriteChartRootReferences(
                        chartPart.ChartSpace,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        rewriteUnqualifiedReferences: false);
                }
            }

            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                if (worksheetPart.DrawingsPart == null) {
                    continue;
                }
                bool rewriteUnqualifiedReferences = ReferenceEquals(worksheetPart, _worksheetPart);
                foreach (ExtendedChartPart chartPart in worksheetPart.DrawingsPart.ExtendedChartParts) {
                    RewriteChartRootReferences(
                        chartPart.ChartSpace,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        rewriteUnqualifiedReferences);
                }
            }
        }

        private void RewriteDrawingShapeTextLinks(
            DrawingsPart? drawingsPart,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            bool rewriteUnqualifiedReferences) {
            Xdr.WorksheetDrawing? drawing = drawingsPart?.WorksheetDrawing;
            if (drawing == null) {
                return;
            }

            bool changed = false;
            foreach (Xdr.Shape shape in drawing.Descendants<Xdr.Shape>()) {
                if (shape.TextLink?.Value is not string formula || formula.Length == 0) {
                    continue;
                }

                string rewritten = lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaReferences(
                        formula,
                        firstAffectedRow,
                        lastDeletedRow.Value,
                        rowDelta,
                        Name,
                        rewriteUnqualifiedReferences)
                    : RewriteShiftedFormulaReferences(
                        formula,
                        firstAffectedRow,
                        rowDelta,
                        Name,
                        rewriteUnqualifiedReferences);
                if (!string.Equals(formula, rewritten, StringComparison.Ordinal)) {
                    shape.TextLink = rewritten;
                    changed = true;
                }
            }

            if (changed) {
                drawing.Save();
            }
        }

        private void RewriteChartRootReferences(
            OpenXmlPartRootElement? chartRoot,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            bool rewriteUnqualifiedReferences) {
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
                    rewriteUnqualifiedReferences);
                changed |= formulaChanged;
                if (formulaChanged) {
                    InvalidateChartFormulaCache(formula);
                }
            }
            if (changed) {
                chartRoot.Save();
            }
        }

        private static bool InvalidateChartFormulaCache(OpenXmlLeafTextElement formula) {
            OpenXmlElement? reference = formula.Parent;
            if (reference == null) {
                return false;
            }
            bool changed = false;
            foreach (OpenXmlElement cache in reference.ChildElements
                .Where(element => element.LocalName.EndsWith("Cache", StringComparison.OrdinalIgnoreCase))
                .ToList()) {
                cache.Remove();
                changed = true;
            }
            if (string.Equals(reference.LocalName, "numDim", StringComparison.Ordinal)
                || string.Equals(reference.LocalName, "strDim", StringComparison.Ordinal)) {
                foreach (OpenXmlElement level in reference.ChildElements
                    .Where(element => string.Equals(element.LocalName, "lvl", StringComparison.Ordinal))
                    .ToList()) {
                    level.Remove();
                    changed = true;
                }
            }
            return changed;
        }

        private void InvalidateWorkbookChartCaches() {
            IEnumerable<OpenXmlPartRootElement> chartRoots =
                WorkbookPartRoot.WorksheetParts
                    .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
                    .Select(part => part.ChartSpace)
                    .Where(root => root != null)
                    .Cast<OpenXmlPartRootElement>()
                .Concat(
                    WorkbookPartRoot.WorksheetParts
                        .SelectMany(part => part.DrawingsPart?.ExtendedChartParts ?? Enumerable.Empty<ExtendedChartPart>())
                        .Select(part => part.ChartSpace)
                        .Where(root => root != null)
                        .Cast<OpenXmlPartRootElement>())
                .Concat(
                    WorkbookPartRoot.ChartsheetParts
                        .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
                        .Select(part => part.ChartSpace)
                        .Where(root => root != null)
                        .Cast<OpenXmlPartRootElement>())
                .Concat(
                    WorkbookPartRoot.ChartsheetParts
                        .SelectMany(part => part.DrawingsPart?.ExtendedChartParts ?? Enumerable.Empty<ExtendedChartPart>())
                        .Select(part => part.ChartSpace)
                        .Where(root => root != null)
                        .Cast<OpenXmlPartRootElement>());

            foreach (OpenXmlPartRootElement chartRoot in chartRoots.Distinct()) {
                bool changed = false;
                foreach (OpenXmlLeafTextElement formula in chartRoot.Descendants<OpenXmlLeafTextElement>()
                    .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                    changed |= InvalidateChartFormulaCache(formula);
                }
                if (changed) {
                    chartRoot.Save();
                }
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
