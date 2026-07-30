using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

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
                PreflightPendingDirectCellsForRowInsertion(firstRow, count);
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
            int rows = 0;
            int formulas = 0;
            int definedNames = 0;
            int validation = 0;
            int conditionalFormatting = 0;
            int mergedCells = 0;
            int hyperlinks = 0;
            int dataConsolidation = 0;
            int namedSheetViews = 0;
            int protectedRanges = 0;
            int worksheetRangeMetadata = 0;
            int queryTableSorts = 0;
            int webPublishItems = 0;
            int tables = 0;
            int sparklines = 0;
            int drawings = 0;
            int pivots = 0;
            int comments = 0;
            int connectionParameters = 0;
            var targetCellCoordinates = new HashSet<long>();
            var externalValidationImpacts = new HashSet<OpenXmlElement>();
            var externalConditionalFormattingImpacts = new HashSet<OpenXmlElement>();
            var externalSparklineImpacts = new HashSet<OpenXmlElement>();
            var pendingOwnerCells = new Dictionary<long, object?>();
            var pendingDirectCells = new List<(ExcelSheet Owner, int Row, int Column, object? Value)>();
            ExcelSheet? pendingOwner = _excelDocument.PendingDirectCellValueSheet;
            var inspectedDrawingRoots = new HashSet<OpenXmlPartRootElement>();
            var drawingOwners = new List<(DrawingsPart Part, bool RewriteUnqualifiedReferences)>();
            int mutatedSheetIndex = -1;
            int sheetIndex = 0;

            if (pendingOwner?._pendingCellValueDirectSaveBuffer != null) {
                foreach ((int Row, int Column, object? Value) pending in
                    pendingOwner._pendingCellValueDirectSaveBuffer.EnumerateWrittenCells()) {
                    budget.Consume();
                    long coordinate = ((long)pending.Row << 32) | (uint)pending.Column;
                    pendingOwnerCells[coordinate] = pending.Value;
                    pendingDirectCells.Add((pendingOwner, pending.Row, pending.Column, pending.Value));
                }
            }

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
                var worksheetElements = new List<OpenXmlElement>();
                foreach (OpenXmlElement element in worksheetPart.Worksheet.Descendants()) {
                    budget.Consume();
                    worksheetElements.Add(element);
                }
                ExcelSheet inspectedSheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(
                        _excelDocument,
                        _spreadSheetDocument,
                        sheetElement,
                        registerSheetWrapper: false);
                IReadOnlyDictionary<Cell, (int Row, int Column)> effectiveCoordinates =
                    inspectedSheet.BuildEffectiveCellCoordinates();
                IReadOnlyDictionary<uint, SharedFormulaDefinition> sharedFormulaDefinitions =
                    inspectedSheet.BuildSharedFormulaDefinitions(effectiveCoordinates);
                bool isPendingOwner = pendingOwner != null
                    && ReferenceEquals(worksheetPart, pendingOwner._worksheetPart);
                if (rewriteUnqualified) {
                    rows += CountAffectedRowRecords(worksheetPart.Worksheet, firstRow);
                }
                foreach (OpenXmlElement element in worksheetElements) {
                    if (element is Cell cell) {
                        if (!effectiveCoordinates.TryGetValue(cell, out var effectiveCoordinate)) {
                            continue;
                        }

                        long coordinate = ((long)effectiveCoordinate.Row << 32) | (uint)effectiveCoordinate.Column;
                        bool pendingValueIsAuthoritative = isPendingOwner
                            && pendingOwnerCells.ContainsKey(coordinate);
                        if (rewriteUnqualified
                            && targetCellCoordinates.Add(coordinate)
                            && effectiveCoordinate.Row >= firstRow) {
                            cells++;
                        }

                        if (!pendingValueIsAuthoritative
                            && cell.CellFormula is CellFormula cellFormula) {
                            bool formulaImpactRecorded = false;
                            if (FormulaChangesForPlan(
                                    inspectedSheet.ResolveCellFormulaText(
                                        cell,
                                        sharedFormulaDefinitions,
                                        effectiveCoordinates),
                                    kind,
                                    firstRow,
                                    lastRow,
                                    count,
                                    rewriteUnqualified)) {
                                formulas++;
                                formulaImpactRecorded = true;
                            }
                            if (rewriteUnqualified
                                && cellFormula.FormulaType?.Value != CellFormulaValues.Shared
                                && ReferenceListChangesForPlan(
                                    cellFormula.Reference?.Value,
                                    kind,
                                    firstRow,
                                    lastRow,
                                    count)) {
                                formulas++;
                                formulaImpactRecorded = true;
                            }
                            if (rewriteUnqualified
                                && cellFormula.FormulaType?.Value == CellFormulaValues.DataTable) {
                                if (ReferenceListChangesForPlan(
                                        cellFormula.R1?.Value,
                                        kind,
                                        firstRow,
                                        lastRow,
                                        count)) {
                                    formulas++;
                                    formulaImpactRecorded = true;
                                }
                                if (ReferenceListChangesForPlan(
                                        cellFormula.R2?.Value,
                                        kind,
                                        firstRow,
                                        lastRow,
                                        count)) {
                                    formulas++;
                                    formulaImpactRecorded = true;
                                }
                            }
                            if (cellFormula.FormulaType?.Value == CellFormulaValues.Shared
                                && !formulaImpactRecorded) {
                                formulas++;
                                formulaImpactRecorded = true;
                            }
                            if (!formulaImpactRecorded
                                && cellFormula.CalculateCell?.Value != true) {
                                formulas++;
                            }
                        }
                    } else if (element is OpenXmlLeafTextElement formula
                        && element is not DocumentFormat.OpenXml.Spreadsheet.CellFormula
                        && IsStructuralFormulaElement(formula)
                        && (!rewriteUnqualified || !UsesAnchoredTargetFormulaSemantics(formula))
                        && FormulaChangesForPlan(
                            formula.Text,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualified)) {
                        formulas++;
                        if (!rewriteUnqualified) {
                            ClassifyExternalFormulaPlanImpact(
                                formula,
                                externalValidationImpacts,
                                externalConditionalFormattingImpacts,
                                externalSparklineImpacts);
                        }
                    }
                    if (element is Hyperlink hyperlink) {
                        bool anchorChanges = rewriteUnqualified
                            && ReferenceListChangesForPlan(
                                hyperlink.Reference?.Value,
                                kind,
                                firstRow,
                                lastRow,
                                count);
                        bool internalLocationChanges = string.IsNullOrWhiteSpace(hyperlink.Id?.Value)
                            && FormulaChangesForPlan(
                                hyperlink.Location?.Value,
                                kind,
                                firstRow,
                                lastRow,
                                count,
                                rewriteUnqualified);
                        if (anchorChanges || internalLocationChanges) {
                            hyperlinks++;
                        }
                        if (internalLocationChanges) {
                            formulas++;
                        }
                    }
                    if (!rewriteUnqualified
                        && element is ConditionalFormatValueObject threshold
                        && threshold.Type?.Value == ConditionalFormatValueObjectValues.Formula
                        && FormulaChangesForPlan(
                            threshold.Val?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences: false)) {
                        formulas++;
                        ClassifyExternalFormulaPlanImpact(
                            threshold,
                            externalValidationImpacts,
                            externalConditionalFormattingImpacts,
                            externalSparklineImpacts);
                    }
                    if (element is DataReference source
                        && string.IsNullOrWhiteSpace(source.Id?.Value)
                        && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                        && ReferenceListChangesForPlan(
                            source.Reference?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        dataConsolidation++;
                    }
                    if (!rewriteUnqualified) {
                        continue;
                    }
                    if (element is DataValidation
                        || element is X14.DataValidation
                        || element is ConditionalFormatting
                        || element is X14.ConditionalFormatting) {
                        int metadataFormulaImpacts =
                            CountAnchoredMetadataFormulaPlanImpacts(
                            element,
                            kind,
                            firstRow,
                            lastRow,
                            count);
                        formulas += metadataFormulaImpacts;
                        bool metadataRangeChanges =
                            StructuralMetadataRangeChangesForPlan(
                                element,
                                kind,
                                firstRow,
                                lastRow,
                                count);
                        if (element is DataValidation || element is X14.DataValidation) {
                            if (metadataRangeChanges || metadataFormulaImpacts > 0) {
                                validation++;
                            }
                        } else if (metadataRangeChanges || metadataFormulaImpacts > 0) {
                            conditionalFormatting++;
                        }
                    } else if (element is MergeCell merge
                        && merge.Reference?.Value is string mergeReference
                        && TryParseReference(mergeReference, out var mergeBounds)
                        && TryRemapShiftedReferenceRows(
                            mergeBounds,
                            firstRow,
                            kind == ExcelRowMutationKind.Insert ? count : -count,
                            kind == ExcelRowMutationKind.Delete ? lastRow : (int?)null,
                            out _)) {
                        mergedCells++;
                    } else if (element is DocumentFormat.OpenXml.Office2010.Excel.Sparkline sparkline
                        && SparklineChangesForPlan(
                            sparkline,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        sparklines++;
                    } else if (element is ProtectedRange protectedRange
                        && ReferenceListChangesForPlan(
                            protectedRange.SequenceOfReferences?.InnerText,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        protectedRanges++;
                    }
                }

                if (rewriteUnqualified) {
                    worksheetRangeMetadata += CountWorksheetRangeMetadataPlanImpacts(
                        worksheetPart.Worksheet,
                        kind,
                        firstRow,
                        lastRow,
                        count);
                }

                if (worksheetPart.DrawingsPart != null) {
                    drawingOwners.Add((worksheetPart.DrawingsPart, rewriteUnqualified));
                }

                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    int tableFormulaImpacts = CountTableFormulaPlanImpacts(
                        tablePart,
                        rewriteUnqualified,
                        kind,
                        firstRow,
                        lastRow,
                        count,
                        budget);
                    formulas += tableFormulaImpacts;
                    if ((rewriteUnqualified
                            && TableMetadataChangesForPlan(
                                tablePart.Table,
                                kind,
                                firstRow,
                                lastRow,
                                count))
                        || tableFormulaImpacts > 0) {
                        tables++;
                    }
                }
            }

            foreach ((ExcelSheet Owner, int Row, int Column, object? Value) pending in pendingDirectCells) {
                bool ownerIsTarget = ReferenceEquals(pending.Owner._worksheetPart, _worksheetPart);
                long coordinate = ((long)pending.Row << 32) | (uint)pending.Column;
                if (ownerIsTarget
                    && targetCellCoordinates.Add(coordinate)
                    && pending.Row >= firstRow) {
                    cells++;
                }
                if (pending.Value is DirectFormulaCellValue pendingFormula
                    && FormulaChangesForPlan(
                        pendingFormula.Formula,
                        kind,
                        firstRow,
                        lastRow,
                        count,
                        rewriteUnqualifiedReferences: ownerIsTarget)) {
                    formulas++;
                }
            }

            foreach ((DrawingsPart Part, bool RewriteUnqualifiedReferences) drawingOwner in
                drawingOwners.OrderByDescending(owner => owner.RewriteUnqualifiedReferences)) {
                CountDrawingPlanImpacts(
                    drawingOwner.Part,
                    drawingOwner.RewriteUnqualifiedReferences,
                    kind,
                    firstRow,
                    lastRow,
                    count,
                    budget,
                    inspectedDrawingRoots,
                    ref drawings,
                    ref formulas);
            }

            foreach (ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                budget.Consume();
                CountDrawingPlanImpacts(
                    chartsheetPart.DrawingsPart,
                    rewriteUnqualifiedReferences: false,
                    kind,
                    firstRow,
                    lastRow,
                    count,
                    budget,
                    inspectedDrawingRoots,
                    ref drawings,
                    ref formulas);
            }

            foreach (PivotTablePart pivotPart in _worksheetPart.PivotTableParts) {
                budget.Consume();
                if (ReferenceListChangesForPlan(
                        pivotPart.PivotTableDefinition?.Location?.Reference?.Value,
                        kind,
                        firstRow,
                        lastRow,
                        count)) {
                    pivots++;
                }
            }
            foreach (PivotTableCacheDefinitionPart cachePart in GetWorkbookPivotCacheDefinitionParts()) {
                budget.Consume();
                foreach (OpenXmlElement _ in cachePart.PivotCacheDefinition?.Descendants()
                    ?? Enumerable.Empty<OpenXmlElement>()) {
                    budget.Consume();
                }
                if (PivotSourceChangesForPlan(cachePart, firstRow, lastRow, count, kind)) {
                    pivots++;
                }
            }

            var queryConnectionIds = new HashSet<uint>();
            foreach (QueryTablePart queryPart in _worksheetPart.QueryTableParts) {
                budget.Consume();
                foreach (OpenXmlElement _ in queryPart.QueryTable?.Descendants()
                    ?? Enumerable.Empty<OpenXmlElement>()) {
                    budget.Consume();
                }
                queryTableSorts += CountQueryTableSortPlanImpacts(
                    queryPart.QueryTable,
                    kind,
                    firstRow,
                    lastRow,
                    count);
                if (queryPart.QueryTable?.ConnectionId?.Value is uint connectionId) {
                    queryConnectionIds.Add(connectionId);
                }
            }
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections != null) {
                budget.Consume();
                foreach (Connection connection in connections.Elements<Connection>()) {
                    budget.Consume();
                    bool isTargetConnection = connection.Id?.Value is uint connectionId
                        && queryConnectionIds.Contains(connectionId);
                    foreach (OpenXmlElement element in connection.Descendants()) {
                        budget.Consume();
                        if (isTargetConnection
                            && element is Parameter parameter
                            && ReferenceListChangesForPlan(
                                parameter.Cell?.Value,
                                kind,
                                firstRow,
                                lastRow,
                                count)) {
                            connectionParameters++;
                        }
                    }
                }
            }

            comments += CountCommentPlanImpacts(
                kind,
                firstRow,
                lastRow,
                count,
                budget);

            namedSheetViews += CountNamedSheetViewPlanImpacts(
                kind,
                firstRow,
                lastRow,
                count,
                budget);

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
                if (element is WebPublishItem item
                    && item.SourceType?.Value == WebSourceValues.Range
                    && string.Equals(item.SourceObject?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    && ReferenceListChangesForPlan(
                        item.SourceRef?.Value,
                        kind,
                        firstRow,
                        lastRow,
                        count)) {
                    webPublishItems++;
                }
            }
            validation += externalValidationImpacts.Count;
            conditionalFormatting += externalConditionalFormattingImpacts.Count;
            sparklines += externalSparklineImpacts.Count;
            AddImpact(
                impacts,
                "worksheet-cells",
                cells,
                "Cells at or below the structural boundary can move or be removed.");
            AddImpact(
                impacts,
                "worksheet-rows",
                rows,
                "Worksheet row records at or below the structural boundary can move or be removed.");
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
                tables,
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
                "data-consolidation",
                dataConsolidation,
                "Workbook data-consolidation sources that target this worksheet are checked and remapped.");
            AddImpact(
                impacts,
                "named-sheet-views",
                namedSheetViews,
                "Named-sheet-view filter ranges are checked and remapped.");
            AddImpact(
                impacts,
                "protected-ranges",
                protectedRanges,
                "Editable protected-range metadata is checked and remapped.");
            AddImpact(
                impacts,
                "worksheet-range-metadata",
                worksheetRangeMetadata,
                "Worksheet filters, views, scenarios, watches, errors, and page ranges are checked and remapped.");
            AddImpact(
                impacts,
                "query-table-sorts",
                queryTableSorts,
                "Query-table sort ranges are checked and remapped.");
            AddImpact(
                impacts,
                "web-publish",
                webPublishItems,
                "Workbook web-publish source ranges are checked and remapped.");
            AddImpact(
                impacts,
                "drawings",
                drawings,
                "Drawing anchors, shapes, and embedded chart locations are checked and remapped.");
            AddImpact(
                impacts,
                "pivots",
                pivots,
                "Pivot output and source boundaries are validated before mutation.");
            AddImpact(
                impacts,
                "comments",
                comments,
                "Legacy and threaded comment anchors are checked and remapped.");
            AddImpact(
                impacts,
                "connection-parameters",
                connectionParameters,
                "Cell-backed query connection parameters are checked and remapped.");
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
            if (_excelDocument.HasUnmaterializedDirectDataSetRows) {
                throw new InvalidOperationException(
                    "A non-mutating structural plan cannot inspect pending deferred or preserved fast-save worksheet rows. " +
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

        private static bool ReferenceListChangesForPlan(
            string? reference,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            return !string.IsNullOrWhiteSpace(reference)
                && TryRemapShiftedReferenceListRows(
                    reference!,
                    firstRow,
                    kind == ExcelRowMutationKind.Insert ? count : -count,
                    kind == ExcelRowMutationKind.Delete ? lastRow : null,
                    out _);
        }

        private static bool StructuralMetadataRangeChangesForPlan(
            OpenXmlElement metadata,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            string? references = metadata switch {
                DataValidation validation =>
                    validation.SequenceOfReferences?.InnerText,
                X14.DataValidation validation =>
                    validation.ReferenceSequence?.Text,
                ConditionalFormatting formatting =>
                    formatting.SequenceOfReferences?.InnerText,
                X14.ConditionalFormatting formatting =>
                    formatting.GetFirstChild<
                        DocumentFormat.OpenXml.Office.Excel.ReferenceSequence>()?.Text,
                _ => null
            };
            return ReferenceListChangesForPlan(
                references,
                kind,
                firstRow,
                lastRow,
                count);
        }

        private void PreflightPendingDirectCellsForRowInsertion(int firstRow, int count) {
            ExcelSheet? pendingOwner = _excelDocument.PendingDirectCellValueSheet;
            if (pendingOwner?._pendingCellValueDirectSaveBuffer == null) {
                return;
            }

            bool ownerIsTarget = ReferenceEquals(pendingOwner._worksheetPart, _worksheetPart);
            foreach ((int Row, int Column, object? Value) pending in
                pendingOwner._pendingCellValueDirectSaveBuffer.EnumerateWrittenCells()) {
                if (ownerIsTarget
                    && pending.Row >= firstRow
                    && (long)pending.Row + count > A1.MaxRows) {
                    throw new InvalidOperationException(
                        "Inserting rows would move a pending worksheet cell beyond Excel's row limit.");
                }
                if (pending.Value is DirectFormulaCellValue formula) {
                    ThrowIfFormulaReferenceOverflows(
                        formula.Formula,
                        firstRow,
                        count,
                        rewriteUnqualifiedReferences: ownerIsTarget);
                }
            }
        }

        private bool PivotSourceChangesForPlan(
            PivotTableCacheDefinitionPart cachePart,
            int firstRow,
            int lastRow,
            int count,
            ExcelRowMutationKind kind) {
            WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
            if (source != null
                && string.IsNullOrWhiteSpace(source.Id?.Value)
                && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                && source.Reference?.Value is string sourceReference
                && FormulaChangesForPlan(
                    sourceReference,
                    kind,
                    firstRow,
                    lastRow,
                    count,
                    rewriteUnqualifiedReferences: true)) {
                return true;
            }

            if (string.IsNullOrWhiteSpace(source?.Id?.Value)
                && source?.Name?.Value is string sourceName
                && IsNamedPivotSourceAffected(sourceName, firstRow)) {
                return true;
            }

            return cachePart.PivotCacheDefinition?.CacheSource?.Consolidation?.RangeSets?
                .Elements<RangeSet>()
                .Any(rangeSet =>
                    string.IsNullOrWhiteSpace(rangeSet.Id?.Value)
                    && string.Equals(rangeSet.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    && FormulaChangesForPlan(
                        rangeSet.Reference?.Value,
                        kind,
                        firstRow,
                        lastRow,
                        count,
                        rewriteUnqualifiedReferences: true)) == true;
        }

        private void CountDrawingPlanImpacts(
            DrawingsPart? drawingsPart,
            bool rewriteUnqualifiedReferences,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count,
            MutationPlanScanBudget budget,
            ISet<OpenXmlPartRootElement> inspectedRoots,
            ref int drawings,
            ref int formulas) {
            if (drawingsPart == null) {
                return;
            }

            if (drawingsPart.WorksheetDrawing != null
                && inspectedRoots.Add(drawingsPart.WorksheetDrawing)) {
                budget.Consume();
                var changedDrawingItems = new HashSet<OpenXmlElement>();
                foreach (OpenXmlElement element in drawingsPart.WorksheetDrawing.Descendants()) {
                    budget.Consume();
                    if (rewriteUnqualifiedReferences
                        && (element is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor
                            || element is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor)
                        && DrawingAnchorChangesForPlan(
                            element,
                            kind,
                            firstRow,
                            lastRow,
                            count)) {
                        changedDrawingItems.Add(element);
                    }
                    if (element is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape
                        && FormulaChangesForPlan(
                            shape.TextLink?.Value,
                            kind,
                            firstRow,
                            lastRow,
                            count,
                            rewriteUnqualifiedReferences)) {
                        OpenXmlElement changedItem = shape.Ancestors()
                            .FirstOrDefault(ancestor =>
                                ancestor is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor
                                || ancestor is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor)
                            ?? shape;
                        changedDrawingItems.Add(changedItem);
                        formulas++;
                    }
                }
                drawings += changedDrawingItems.Count;
            }

            foreach (OpenXmlPartRootElement? chartRoot in drawingsPart.ChartParts
                .Select(part => part.ChartSpace)
                .Cast<OpenXmlPartRootElement?>()
                .Concat(drawingsPart.ExtendedChartParts.Select(part => part.ChartSpace))) {
                if (chartRoot == null || !inspectedRoots.Add(chartRoot)) {
                    continue;
                }

                budget.Consume();
                bool chartChanges = false;
                foreach (OpenXmlElement element in chartRoot.Descendants()) {
                    budget.Consume();
                    if (element is OpenXmlLeafTextElement formula
                        && string.Equals(formula.LocalName, "f", StringComparison.Ordinal)) {
                        if (FormulaChangesForPlan(
                                formula.Text,
                                kind,
                                firstRow,
                                lastRow,
                                count,
                                rewriteUnqualifiedReferences)) {
                            formulas++;
                            chartChanges = true;
                        }
                        if (ChartFormulaCacheWillBeInvalidated(formula)) {
                            chartChanges = true;
                        }
                    }
                }
                if (chartChanges) {
                    drawings++;
                }
            }
        }

        private bool DrawingAnchorChangesForPlan(
            OpenXmlElement anchor,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) {
            int rowDelta = kind == ExcelRowMutationKind.Insert ? count : -count;
            int? lastDeletedRow = kind == ExcelRowMutationKind.Delete ? lastRow : null;

            if (anchor is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor) {
                return DrawingMarkerChangesForPlan(
                    oneCellAnchor.FromMarker,
                    firstRow,
                    rowDelta,
                    lastDeletedRow);
            }

            if (anchor is not DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor) {
                return false;
            }

            DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues placement =
                twoCellAnchor.EditAs?.Value
                ?? DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.TwoCell;
            if (placement == DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.Absolute) {
                return false;
            }
            if (placement == DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.OneCell) {
                return DrawingMarkerChangesForPlan(
                    twoCellAnchor.FromMarker,
                    firstRow,
                    rowDelta,
                    lastDeletedRow);
            }

            return TwoCellDrawingAnchorChangesForPlan(
                twoCellAnchor,
                firstRow,
                rowDelta,
                lastDeletedRow);
        }

        private static bool DrawingMarkerChangesForPlan(
            DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType? marker,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            if (marker?.RowId?.Text is not string rowText
                || !int.TryParse(
                    rowText,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out int zeroBasedRow)) {
                return false;
            }

            int oneBasedRow = zeroBasedRow + 1;
            if (!TryRemapShiftedReferenceRows(
                    (oneBasedRow, 1, oneBasedRow, 1),
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out var remapped)) {
                return false;
            }
            if (remapped == null) {
                return !lastDeletedRow.HasValue
                    || firstAffectedRow - 1 != zeroBasedRow;
            }
            return remapped.Value.r1 - 1 != zeroBasedRow;
        }

        private bool SparklineChangesForPlan(
            X14.Sparkline sparkline,
            ExcelRowMutationKind kind,
            int firstRow,
            int lastRow,
            int count) =>
            ReferenceListChangesForPlan(
                sparkline.ReferenceSequence?.Text,
                kind,
                firstRow,
                lastRow,
                count)
            || FormulaChangesForPlan(
                sparkline.Formula?.Text,
                kind,
                firstRow,
                lastRow,
                count,
                rewriteUnqualifiedReferences: true);

        private static bool TwoCellDrawingAnchorChangesForPlan(
            DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor anchor,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            if (rowDelta == 0
                || anchor.FromMarker?.RowId?.Text is not string fromRowText
                || anchor.ToMarker?.RowId?.Text is not string toRowText
                || !int.TryParse(
                    fromRowText,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out int fromZeroBasedRow)
                || !int.TryParse(
                    toRowText,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out int toZeroBasedRow)) {
                return false;
            }

            int firstSpannedRow = fromZeroBasedRow + 1;
            bool toMarkerInsideRow = long.TryParse(
                    anchor.ToMarker?.RowOffset?.Text,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out long toRowOffset)
                && toRowOffset != 0L;
            int lastSpannedRow = toZeroBasedRow + (toMarkerInsideRow ? 1 : 0);
            if (lastSpannedRow < firstSpannedRow
                || !TryRemapShiftedReferenceRows(
                    (firstSpannedRow, 1, lastSpannedRow, 1),
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out var remapped)) {
                return false;
            }
            if (remapped == null) {
                return !lastDeletedRow.HasValue
                    || firstAffectedRow - 1 != fromZeroBasedRow
                    || firstAffectedRow - 1 != toZeroBasedRow;
            }

            int remappedFromZeroBasedRow = remapped.Value.r1 - 1;
            int remappedToZeroBasedRow = remapped.Value.r2 - (toMarkerInsideRow ? 1 : 0);
            return remappedFromZeroBasedRow != fromZeroBasedRow
                || remappedToZeroBasedRow != toZeroBasedRow;
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
