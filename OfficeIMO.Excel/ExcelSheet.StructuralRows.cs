using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xnsv = DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Inserts blank worksheet rows immediately before <paramref name="firstRow"/> and shifts existing rows down.
        /// Formulas, defined names, tables, merges, comments, hyperlinks, validations, conditional formatting,
        /// sparklines, drawings, and chart references owned by the workbook are adjusted where supported.
        /// </summary>
        /// <param name="firstRow">The 1-based row before which new rows are inserted.</param>
        /// <param name="count">The number of rows to insert.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the requested rows fall outside the worksheet row limit.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when content or a dependent reference would exceed the worksheet row limit, or when the insertion
        /// would split an array formula, data-table, or PivotTable output range, when the workbook uses R1C1 reference
        /// mode, or when the worksheet contains unsupported form controls, OLE objects, or single-cell XML mappings.
        /// </exception>
        public void InsertRows(int firstRow, int count = 1) {
            ValidateStructuralRowArguments(firstRow, count);
            WriteLockConditional(() => {
                ShiftRowsDown(firstRow, count);
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Deletes worksheet rows starting at <paramref name="firstRow"/> and shifts following rows up.
        /// Formula references to cells fully removed by the operation become <c>#REF!</c>; partially overlapping
        /// ranges are reduced to the surviving cells.
        /// </summary>
        /// <param name="firstRow">The 1-based first row to delete.</param>
        /// <param name="count">The number of rows to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the requested rows fall outside the worksheet row limit.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the deletion would remove an owned table boundary or intersect an array-formula, data-table,
        /// or PivotTable output range that cannot be rebuilt safely, when the workbook uses R1C1 reference mode,
        /// or when the worksheet contains unsupported form controls, OLE objects, or single-cell XML mappings.
        /// </exception>
        public void DeleteRows(int firstRow, int count = 1) {
            ValidateStructuralRowArguments(firstRow, count);
            WriteLockConditional(() => {
                RemoveRowsAndShiftUp(firstRow, count);
                WorksheetRoot.Save();
            });
        }

        private static void ValidateStructuralRowArguments(int firstRow, int count) {
            if (firstRow < 1 || firstRow > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(firstRow), $"Row must be between 1 and {A1.MaxRows}.");
            }
            if (count < 1) {
                throw new ArgumentOutOfRangeException(nameof(count), "Row count must be greater than zero.");
            }
            if ((long)firstRow + count - 1L > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(count), "The requested row block must fit inside the worksheet row limit.");
            }
        }

        private void ShiftRowsDown(int firstRow, int count) {
            if (count <= 0) {
                return;
            }

            ValidateStructuralRowReferenceMode();
            ValidateStructuralRowControlSafety();
            ValidateRowInsertionAgainstArrayFormulas(firstRow);
            ValidateRowInsertionAgainstPivotOutputs(firstRow);
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            uint maxShiftedRow = GetMaximumEffectiveRowIndex(sheetData, firstRow);
            if (maxShiftedRow > 0U && (long)maxShiftedRow + count > A1.MaxRows) {
                throw new InvalidOperationException("Inserting rows would move worksheet content beyond Excel's row limit.");
            }
            ValidateStructuralRowReferenceCapacity(firstRow, count);
            MaterializeWorkbookSharedFormulasForStructuralEdit();

            if (sheetData != null) {
                NormalizeImplicitRowIndices(sheetData);
                foreach (Row row in sheetData.Elements<Row>()
                    .Where(item => item.RowIndex?.Value >= (uint)firstRow)
                    .OrderByDescending(item => item.RowIndex?.Value ?? 0U)
                    .ToList()) {
                    int newRowIndex = checked((int)(row.RowIndex!.Value + (uint)count));
                    row.RowIndex = (uint)newRowIndex;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        if (cell.CellReference?.Value is not string reference || reference.Length == 0) {
                            continue;
                        }

                        int column = GetColumnIndex(reference);
                        if (column > 0) {
                            cell.CellReference = BuildCellReference(newRowIndex, column);
                        }
                    }
                }
            }

            RewriteWorksheetFormulaReferences(firstRow, count);
            RemapShiftedRowMetadata(firstRow, count);
            ShiftMergeCellsRows(firstRow, count);
            InvalidateStructuralFormulaResults();
            ResetStructuralMutationCaches();
        }

        private void RemoveRowsAndShiftUp(int firstRow, int count) {
            if (count <= 0) {
                return;
            }

            int lastRemovedRow = firstRow + count - 1;
            ValidateStructuralRowReferenceMode();
            ValidateStructuralRowControlSafety();
            ValidateRowDeletionAgainstOwnedRanges(firstRow, lastRemovedRow);
            ValidateWorkbookSharedFormulasForStructuralEdit();
            MaterializeWorkbookSharedFormulasForStructuralEdit();
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData != null) {
                NormalizeImplicitRowIndices(sheetData);
                foreach (Row row in sheetData.Elements<Row>().ToList()) {
                    int rowIndex = checked((int)row.RowIndex!.Value);
                    if (rowIndex >= firstRow && rowIndex <= lastRemovedRow) {
                        row.Remove();
                        continue;
                    }

                    if (rowIndex > lastRemovedRow) {
                        int newRowIndex = rowIndex - count;
                        row.RowIndex = (uint)newRowIndex;
                        foreach (Cell cell in row.Elements<Cell>()) {
                            if (cell.CellReference?.Value is not string reference || reference.Length == 0) {
                                continue;
                            }

                            int column = GetColumnIndex(reference);
                            if (column > 0) {
                                cell.CellReference = BuildCellReference(newRowIndex, column);
                            }
                        }
                    }
                }
            }

            RewriteDeletedWorksheetFormulaReferences(firstRow, lastRemovedRow, -count);
            RemapDeletedRowMetadata(firstRow, lastRemovedRow, -count);
            ShiftMergeCellsRows(firstRow, -count, lastRemovedRow);
            InvalidateStructuralFormulaResults();
            ResetStructuralMutationCaches();
        }

        private void ValidateRowInsertionAgainstArrayFormulas(int firstRow) {
            foreach (CellFormula formula in WorksheetRoot.Descendants<CellFormula>()
                .Where(item => item.FormulaType?.Value == CellFormulaValues.Array
                    || item.FormulaType?.Value == CellFormulaValues.DataTable)) {
                if (formula.Reference?.Value is not string reference
                    || !A1.TryParseRange(
                        reference.Replace("$", string.Empty),
                        out int arrayFirstRow,
                        out _,
                        out int arrayLastRow,
                        out _)) {
                    continue;
                }

                if (firstRow > arrayFirstRow && firstRow <= arrayLastRow) {
                    string formulaKind = formula.FormulaType?.Value == CellFormulaValues.DataTable
                        ? "data-table"
                        : "array formula";
                    throw new InvalidOperationException(
                        $"Cannot insert rows through {formulaKind} range '{reference}'. Insert before or after the complete range.");
                }
            }
        }

        private void ValidateRowDeletionAgainstOwnedRanges(int firstDeletedRow, int lastDeletedRow) {
            foreach (CellFormula formula in WorksheetRoot.Descendants<CellFormula>()
                .Where(item => item.FormulaType?.Value == CellFormulaValues.Array
                    || item.FormulaType?.Value == CellFormulaValues.DataTable)) {
                if (formula.Reference?.Value is not string reference
                    || !A1.TryParseRange(
                        reference.Replace("$", string.Empty),
                        out int ownedFirstRow,
                        out _,
                        out int ownedLastRow,
                        out _)) {
                    continue;
                }

                bool deletesOwner = ownedFirstRow >= firstDeletedRow && ownedFirstRow <= lastDeletedRow;
                bool deletesWholeRange = firstDeletedRow <= ownedFirstRow && lastDeletedRow >= ownedLastRow;
                if (deletesOwner && !deletesWholeRange) {
                    string formulaKind = formula.FormulaType?.Value == CellFormulaValues.DataTable
                        ? "data-table"
                        : "array formula";
                    throw new InvalidOperationException(
                        $"Cannot delete the owner row of {formulaKind} range '{reference}' while part of the range survives.");
                }
            }

            foreach (var tableDefinitionPart in _worksheetPart.TableDefinitionParts) {
                Table? table = tableDefinitionPart.Table;
                if (table?.Reference?.Value is not string reference
                    || !A1.TryParseRange(
                        reference.Replace("$", string.Empty),
                        out int tableFirstRow,
                        out _,
                        out int tableLastRow,
                        out _)) {
                    continue;
                }

                string tableName = table.Name?.Value ?? table.DisplayName?.Value ?? reference;
                if (firstDeletedRow <= tableFirstRow && lastDeletedRow >= tableLastRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete the complete range of table '{tableName}'. Remove the table first.");
                }

                bool hasHeaderRow = (table.HeaderRowCount?.Value ?? 1U) > 0U;
                if (hasHeaderRow && tableFirstRow >= firstDeletedRow && tableFirstRow <= lastDeletedRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete the header row of table '{tableName}'. Remove or resize the table first.");
                }

                bool hasTotalsRow = HasActiveTotalsRow(table);
                if (hasTotalsRow && tableLastRow >= firstDeletedRow && tableLastRow <= lastDeletedRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete the totals row of table '{tableName}'. Disable totals or resize the table first.");
                }
            }

            foreach (var cachePart in GetWorkbookPivotCacheDefinitionParts()) {
                WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (source?.Reference?.Value is string reference
                    && string.IsNullOrWhiteSpace(source.Id?.Value)
                    && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    && A1.TryParseRange(
                        reference.Replace("$", string.Empty),
                        out int sourceFirstRow,
                        out _,
                        out _,
                        out _)
                    && sourceFirstRow >= firstDeletedRow
                    && sourceFirstRow <= lastDeletedRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete the header row of pivot cache source range '{reference}'. Update or remove the pivot source first.");
                }
                if (string.IsNullOrWhiteSpace(source?.Id?.Value)
                    && source?.Name?.Value is string sourceName) {
                    bool resolvedNamedSource = TryResolveDefinedNameRange(
                            sourceName,
                            currentRow: null,
                            out ExcelSheet sourceSheet,
                            out int namedSourceFirstRow,
                            out _,
                            out _,
                            out _);
                    bool deletesResolvedHeader = resolvedNamedSource
                        && (ReferenceEquals(sourceSheet._worksheetPart, _worksheetPart)
                            || string.Equals(sourceSheet.Name, Name, StringComparison.OrdinalIgnoreCase))
                        && namedSourceFirstRow >= firstDeletedRow
                        && namedSourceFirstRow <= lastDeletedRow;
                    if (deletesResolvedHeader
                        || (!resolvedNamedSource && UnresolvedNamedPivotSourceBecomesInvalid(
                            sourceName,
                            firstDeletedRow,
                            lastDeletedRow))) {
                        throw new InvalidOperationException(
                            $"Cannot delete the header row of named pivot cache source '{sourceName}'. Update or remove the pivot source first.");
                    }
                }

                foreach (RangeSet rangeSet in cachePart.PivotCacheDefinition?.CacheSource?
                    .Consolidation?.RangeSets?.Elements<RangeSet>() ?? Enumerable.Empty<RangeSet>()) {
                    if (string.IsNullOrWhiteSpace(rangeSet.Id?.Value)
                        && string.Equals(rangeSet.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                        && rangeSet.Reference?.Value is string consolidationReference
                        && A1.TryParseRange(
                            consolidationReference.Replace("$", string.Empty),
                            out int consolidationFirstRow,
                            out _,
                            out int consolidationLastRow,
                            out _)
                        && firstDeletedRow <= consolidationFirstRow
                        && lastDeletedRow >= consolidationLastRow) {
                        throw new InvalidOperationException(
                            $"Cannot delete the complete consolidation source range '{consolidationReference}' of a pivot cache. Update or remove the pivot source first.");
                    }
                }
            }

            foreach (var pivotPart in _worksheetPart.PivotTableParts) {
                string? reference = pivotPart.PivotTableDefinition?.Location?.Reference?.Value;
                if (reference == null
                    || !A1.TryParseRange(reference.Replace("$", string.Empty), out int pivotFirstRow, out _, out int pivotLastRow, out _)) {
                    continue;
                }

                if (firstDeletedRow <= pivotLastRow && lastDeletedRow >= pivotFirstRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete rows through pivot table output range '{reference}'. Remove or move the pivot table first.");
                }
            }

            ValidateConnectionParameterDeletion(firstDeletedRow, lastDeletedRow);
        }

        private void ValidateRowInsertionAgainstPivotOutputs(int firstRow) {
            foreach (var pivotPart in _worksheetPart.PivotTableParts) {
                string? reference = pivotPart.PivotTableDefinition?.Location?.Reference?.Value;
                if (reference != null
                    && A1.TryParseRange(reference.Replace("$", string.Empty), out int pivotFirstRow, out _, out int pivotLastRow, out _)
                    && firstRow > pivotFirstRow
                    && firstRow <= pivotLastRow) {
                    throw new InvalidOperationException(
                        $"Cannot insert rows through pivot table output range '{reference}'. Insert before or after the pivot table.");
                }
            }
        }

        private bool UnresolvedNamedPivotSourceBecomesInvalid(
            string sourceName,
            int firstDeletedRow,
            int lastDeletedRow) {
            var catalog = new FormulaDefinedNameResolutionCatalog(this);
            int? localSheetIndex = catalog.TryGetSheet(Name, out int sheetIndex, out _)
                ? sheetIndex
                : null;
            if (!catalog.TryGetDefinedName(
                    localSheetIndex,
                    sourceName,
                    allowGlobal: true,
                    out DefinedName definedName,
                    out _)) {
                return false;
            }

            int rowDelta = -(lastDeletedRow - firstDeletedRow + 1);
            string formula = definedName.Text ?? string.Empty;
            if (formula.Length == 0 || formula.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0) {
                return false;
            }

            string rewritten = RewriteDeletedFormulaReferences(
                formula,
                firstDeletedRow,
                lastDeletedRow,
                rowDelta,
                Name);
            return rewritten.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ValidateStructuralRowReferenceMode() {
            if (WorkbookRoot.GetFirstChild<CalculationProperties>()?.ReferenceMode?.Value == ReferenceModeValues.R1C1) {
                throw new InvalidOperationException(
                    "Structural row edits are not supported while the workbook uses R1C1 reference mode. Switch to A1 reference mode first.");
            }
        }

        private static uint GetMaximumEffectiveRowIndex(SheetData? sheetData, int firstRow) {
            uint maximum = 0U;
            uint previous = 0U;
            foreach (Row row in sheetData?.Elements<Row>() ?? Enumerable.Empty<Row>()) {
                uint effective = GetEffectiveRowIndex(row, previous);
                if (effective >= (uint)firstRow && effective > maximum) {
                    maximum = effective;
                }
                previous = effective;
            }
            return maximum;
        }

        private static void NormalizeImplicitRowIndices(SheetData sheetData) {
            uint previous = 0U;
            foreach (Row row in sheetData.Elements<Row>()) {
                uint effective = GetEffectiveRowIndex(row, previous);
                row.RowIndex = effective;
                previous = effective;
            }
        }

        private static uint GetEffectiveRowIndex(Row row, uint previous) {
            if (row.RowIndex?.Value is uint explicitIndex && explicitIndex > 0U) {
                return explicitIndex;
            }

            foreach (Cell cell in row.Elements<Cell>()) {
                if (cell.CellReference?.Value is string reference
                    && TryParseReference(reference, out var bounds)
                    && bounds.r1 > 0) {
                    return (uint)bounds.r1;
                }
            }

            return checked(previous + 1U);
        }

        private void ValidateStructuralRowReferenceCapacity(int firstRow, int count) {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart) {
                    continue;
                }

                bool rewriteUnqualified = ReferenceEquals(worksheetPart, _worksheetPart)
                    || string.Equals(sheetElement.Name?.Value, Name, StringComparison.OrdinalIgnoreCase);
                ExcelSheet formulaSheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement);
                foreach (string formula in formulaSheet.ResolveSharedFormulaTextsForStructuralValidation()) {
                    ThrowIfFormulaReferenceOverflows(formula, firstRow, count, rewriteUnqualified);
                }

                Worksheet? worksheet = worksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }
                foreach (OpenXmlLeafTextElement formula in worksheet
                    .Descendants<OpenXmlLeafTextElement>()
                    .Where(element => element is CellFormula
                        || element is Formula
                        || element is Formula1
                        || element is Formula2
                        || string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                    ThrowIfFormulaReferenceOverflows(formula.Text, firstRow, count, rewriteUnqualified);
                }
                foreach (ConditionalFormatValueObject threshold in worksheet
                    .Descendants<ConditionalFormatValueObject>()
                    .Where(item => item.Type?.Value == ConditionalFormatValueObjectValues.Formula)) {
                    ThrowIfFormulaReferenceOverflows(
                        threshold.Val?.Value,
                        firstRow,
                        count,
                        rewriteUnqualified);
                }

                foreach (Hyperlink hyperlink in worksheet.Descendants<Hyperlink>()) {
                    if (string.IsNullOrWhiteSpace(hyperlink.Id?.Value)) {
                        ThrowIfFormulaReferenceOverflows(hyperlink.Location?.Value, firstRow, count, rewriteUnqualified);
                    }
                }

                foreach (DataReference source in worksheet.Descendants<DataReference>()) {
                    if (string.IsNullOrWhiteSpace(source.Id?.Value)
                        && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                        ValidateReferenceListDoesNotOverflow(source.Reference?.Value, firstRow, count);
                    }
                }

                foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                    if (tablePart.Table == null) {
                        continue;
                    }
                    foreach (OpenXmlLeafTextElement formula in tablePart.Table.Descendants<OpenXmlLeafTextElement>()
                        .Where(element => element is CalculatedColumnFormula || element is TotalsRowFormula)) {
                        ThrowIfFormulaReferenceOverflows(formula.Text, firstRow, count, rewriteUnqualified);
                    }
                }

                ValidateChartFormulaCapacity(
                    worksheetPart.DrawingsPart,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences: ReferenceEquals(worksheetPart, _worksheetPart));
                ValidateDrawingShapeTextLinkCapacity(
                    worksheetPart.DrawingsPart,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences: ReferenceEquals(worksheetPart, _worksheetPart));
            }

            ValidatePivotReferenceCapacity(firstRow, count);
            ValidateConnectionParameterCapacity(firstRow, count);

            foreach (DocumentFormat.OpenXml.Packaging.ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                ValidateChartFormulaCapacity(
                    chartsheetPart.DrawingsPart,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences: false);
                ValidateDrawingShapeTextLinkCapacity(
                    chartsheetPart.DrawingsPart,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences: false);
            }

            List<Sheet> sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int mutatedSheetIndex = sheets.FindIndex(sheet =>
                string.Equals(sheet.Name?.Value, Name, StringComparison.OrdinalIgnoreCase));
            foreach (DefinedName name in WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
                ThrowIfFormulaReferenceOverflows(
                    name.Text,
                    firstRow,
                    count,
                    mutatedSheetIndex >= 0 && name.LocalSheetId?.Value == (uint)mutatedSheetIndex);
            }

            ValidateReferenceAttributesDoNotOverflow(WorksheetRoot, firstRow, count);
            foreach (CellFormula dataTableFormula in WorksheetRoot.Descendants<CellFormula>()
                .Where(formula => formula.FormulaType?.Value == CellFormulaValues.DataTable)) {
                ValidateReferenceListDoesNotOverflow(dataTableFormula.R1?.Value, firstRow, count);
                ValidateReferenceListDoesNotOverflow(dataTableFormula.R2?.Value, firstRow, count);
            }
            foreach (CellWatch cellWatch in WorksheetRoot.Descendants<CellWatch>()) {
                ValidateReferenceListDoesNotOverflow(cellWatch.CellReference?.Value, firstRow, count);
            }
            foreach (InputCells input in WorksheetRoot.Descendants<InputCells>()) {
                ValidateReferenceListDoesNotOverflow(input.CellReference?.Value, firstRow, count);
            }
            foreach (Selection selection in WorksheetRoot.Descendants<Selection>()) {
                ValidateReferenceListDoesNotOverflow(selection.ActiveCell?.Value, firstRow, count);
            }
            foreach (SheetView view in WorksheetRoot.Descendants<SheetView>()) {
                ValidateReferenceListDoesNotOverflow(view.TopLeftCell?.Value, firstRow, count);
            }
            foreach (CustomSheetView view in WorksheetRoot.Descendants<CustomSheetView>()) {
                ValidateReferenceListDoesNotOverflow(view.TopLeftCell?.Value, firstRow, count);
            }
            foreach (Pane pane in WorksheetRoot.Descendants<Pane>()) {
                ValidateReferenceListDoesNotOverflow(pane.TopLeftCell?.Value, firstRow, count);
            }
            foreach (NamedSheetViewsPart part in _worksheetPart.NamedSheetViewsParts) {
                foreach (Xnsv.NsvFilter filter
                    in part.NamedSheetViews?.Descendants<Xnsv.NsvFilter>()
                    ?? Enumerable.Empty<Xnsv.NsvFilter>()) {
                    ValidateReferenceListDoesNotOverflow(filter.Ref?.Value, firstRow, count);
                }
            }
            foreach (QueryTablePart part in _worksheetPart.QueryTableParts) {
                if (part.QueryTable != null) {
                    ValidateReferenceAttributesDoNotOverflow(part.QueryTable, firstRow, count);
                }
            }
            foreach (OpenXmlElement smartTag in WorksheetRoot.Descendants()
                .Where(element => string.Equals(element.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase))) {
                OpenXmlAttribute referenceAttribute = smartTag.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "r", StringComparison.OrdinalIgnoreCase));
                ValidateReferenceListDoesNotOverflow(referenceAttribute.Value, firstRow, count);
            }
            foreach (WebPublishItem item in WorkbookRoot.Descendants<WebPublishItem>()) {
                if (item.SourceType?.Value == WebSourceValues.Range
                    && string.Equals(item.SourceObject?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                    ValidateReferenceListDoesNotOverflow(item.SourceRef?.Value, firstRow, count);
                }
            }
            foreach (var tablePart in _worksheetPart.TableDefinitionParts) {
                if (tablePart.Table != null) {
                    ValidateReferenceAttributesDoNotOverflow(tablePart.Table, firstRow, count);
                }
            }
            foreach (var commentsPart in _worksheetPart.WorksheetCommentsPart == null
                ? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.WorksheetCommentsPart>()
                : new[] { _worksheetPart.WorksheetCommentsPart }) {
                if (commentsPart.Comments != null) {
                    ValidateReferenceAttributesDoNotOverflow(commentsPart.Comments, firstRow, count);
                }
            }
            foreach (var threadedPart in _worksheetPart.WorksheetThreadedCommentsParts) {
                if (threadedPart.ThreadedComments != null) {
                    ValidateReferenceAttributesDoNotOverflow(threadedPart.ThreadedComments, firstRow, count);
                }
            }
            foreach (OpenXmlLeafTextElement reference in WorksheetRoot.Descendants<OpenXmlLeafTextElement>()
                .Where(element => string.Equals(element.LocalName, "sqref", StringComparison.OrdinalIgnoreCase))) {
                ValidateReferenceListDoesNotOverflow(reference.Text, firstRow, count);
            }

            foreach (Break pageBreak in WorksheetRoot.Descendants<RowBreaks>().SelectMany(rowBreaks => rowBreaks.Elements<Break>())) {
                if (pageBreak.Id?.Value is uint row && row >= firstRow && (long)row + count > A1.MaxRows) {
                    throw new InvalidOperationException("Inserting rows would move a page break beyond Excel's row limit.");
                }
            }

            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType marker in
                _worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType>()
                    .Where(candidate => DrawingMarkerMovesOnInsertion(candidate, firstRow))
                ?? Enumerable.Empty<DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType>()) {
                if (int.TryParse(marker.RowId?.Text, out int zeroBasedRow)
                    && zeroBasedRow + 1 >= firstRow
                    && (long)zeroBasedRow
                        + (IsZeroOffsetDrawingBoundary(marker) ? 0L : 1L)
                        + count > A1.MaxRows) {
                    throw new InvalidOperationException("Inserting rows would move a drawing anchor beyond Excel's row limit.");
                }
            }

            ValidateCommentVmlAnchorCapacity(firstRow, count);
        }

        private void ValidatePivotReferenceCapacity(int firstRow, int count) {
            foreach (DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart cachePart
                in GetWorkbookPivotCacheDefinitionParts()) {
                WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (source != null
                    && string.IsNullOrWhiteSpace(source.Id?.Value)
                    && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                    ValidateReferenceListDoesNotOverflow(source.Reference?.Value, firstRow, count);
                }
                foreach (RangeSet rangeSet in cachePart.PivotCacheDefinition?.CacheSource?
                    .Consolidation?.RangeSets?.Elements<RangeSet>() ?? Enumerable.Empty<RangeSet>()) {
                    if (string.IsNullOrWhiteSpace(rangeSet.Id?.Value)
                        && string.Equals(rangeSet.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                        ValidateReferenceListDoesNotOverflow(rangeSet.Reference?.Value, firstRow, count);
                    }
                }
            }

            foreach (DocumentFormat.OpenXml.Packaging.PivotTablePart pivotPart in _worksheetPart.PivotTableParts) {
                ValidateReferenceListDoesNotOverflow(
                    pivotPart.PivotTableDefinition?.Location?.Reference?.Value,
                    firstRow,
                    count);
            }
        }

        private void ValidateConnectionParameterCapacity(int firstRow, int count) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) {
                return;
            }

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            foreach (Connection connection in connections.Elements<Connection>()
                .Where(connection => connection.Id?.Value is uint id && connectionIds.Contains(id))) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    ValidateReferenceListDoesNotOverflow(parameter.Cell?.Value, firstRow, count);
                }
            }
        }

        private void ValidateConnectionParameterDeletion(int firstDeletedRow, int lastDeletedRow) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) {
                return;
            }

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            foreach (Connection connection in connections.Elements<Connection>()
                .Where(connection => connection.Id?.Value is uint id && connectionIds.Contains(id))) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    if (parameter.Cell?.Value is not string reference
                        || !TryParseReference(reference, out var bounds)
                        || bounds.r1 < firstDeletedRow
                        || bounds.r1 > lastDeletedRow) {
                        continue;
                    }

                    throw new InvalidOperationException(
                        $"Cannot delete cell-backed connection parameter reference '{reference}'. Update or remove the parameter first.");
                }
            }
        }

        private void ValidateStructuralRowControlSafety() {
            IEnumerable<VmlDrawingPart> workbookVmlParts =
                WorkbookPartRoot.WorksheetParts.SelectMany(part => part.VmlDrawingParts)
                    .Concat(WorkbookPartRoot.DialogsheetParts.SelectMany(part => part.VmlDrawingParts))
                    .Concat(WorkbookPartRoot.ChartsheetParts.SelectMany(part => part.VmlDrawingParts))
                    .Distinct();
            if (WorkbookPartRoot.WorksheetParts.Any(worksheetPart =>
                    worksheetPart.Worksheet?.Descendants<Controls>().Any() == true
                    || worksheetPart.ControlPropertiesParts.Any())
                || ContainsUnsupportedVmlFormControl(workbookVmlParts)) {
                throw new InvalidOperationException(
                    "Cannot edit rows in a workbook containing form controls because their anchors and cross-sheet links cannot yet be remapped safely.");
            }
            if (WorksheetRoot.Descendants<OleObjects>().Any()
                || _worksheetPart.EmbeddedObjectParts.Any()) {
                throw new InvalidOperationException(
                    "Cannot edit rows on a worksheet containing embedded OLE objects because their VML anchors cannot yet be remapped safely.");
            }
            if (_worksheetPart.SingleCellTablePart != null) {
                throw new InvalidOperationException(
                    "Cannot edit rows on a worksheet containing single-cell XML mappings because their mapped references cannot yet be remapped safely.");
            }
            if (WorkbookPartRoot.MacroSheetParts.Any()
                || WorkbookPartRoot.InternationalMacroSheetParts.Any()) {
                throw new InvalidOperationException(
                    "Cannot edit rows in a workbook containing Excel 4.0 macro sheets because their formulas cannot yet be remapped safely.");
            }
            if (WorkbookPartRoot.WorkbookRevisionHeaderPart != null) {
                throw new InvalidOperationException(
                    "Cannot edit rows while legacy workbook revision tracking is present because revision-log references cannot yet be remapped safely.");
            }
        }

        private bool ContainsUnsupportedVmlFormControl(IEnumerable<VmlDrawingPart> vmlParts) {
            XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
            foreach (VmlDrawingPart vmlPart in vmlParts) {
                XDocument document = LoadOrCreateVmlDocument(vmlPart);
                foreach (XElement clientData in document.Descendants(excelNamespace + "ClientData")) {
                    string? objectType = clientData.Attribute("ObjectType")?.Value;
                    if (!string.Equals(objectType, "Note", StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private void ValidateCommentVmlAnchorCapacity(int firstRow, int count) {
            XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
            foreach (VmlDrawingPart vmlPart in _worksheetPart.VmlDrawingParts) {
                XDocument document = LoadOrCreateVmlDocument(vmlPart);
                foreach (XElement clientData in document.Descendants(excelNamespace + "ClientData")
                    .Where(element => string.Equals(
                        element.Attribute("ObjectType")?.Value,
                        "Note",
                        StringComparison.OrdinalIgnoreCase))) {
                    VmlAnchorPlacement placement = GetVmlAnchorPlacement(clientData, excelNamespace);
                    if (placement == VmlAnchorPlacement.Absolute) {
                        continue;
                    }
                    if (!TryParseVmlAnchor(clientData.Element(excelNamespace + "Anchor"), out int[] values)) {
                        continue;
                    }

                    if (placement == VmlAnchorPlacement.OneCell) {
                        int oneBasedFromRow = values[2] + 1;
                        if (oneBasedFromRow >= firstRow
                            && (long)values[6] + count > A1.MaxRows) {
                            throw new InvalidOperationException(
                                "Inserting rows would move a comment note anchor beyond Excel's row limit.");
                        }
                        continue;
                    }

                    int firstSpannedRow = values[2] + 1;
                    int lastSpannedRow = values[6];
                    if (lastSpannedRow >= firstSpannedRow
                        && TryRemapShiftedReferenceRows(
                            (firstSpannedRow, 1, lastSpannedRow, 1),
                            firstRow,
                            count,
                            lastDeletedRow: null,
                            out var remappedRows)
                        && remappedRows == null) {
                        throw new InvalidOperationException(
                            "Inserting rows would move a comment note anchor beyond Excel's row limit.");
                    }
                }
            }
        }

        private static bool DrawingMarkerMovesOnInsertion(
            DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType marker,
            int firstRow) {
            DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor? anchor =
                marker.Ancestors<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                    .FirstOrDefault();
            if (anchor == null) {
                return true;
            }

            DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues placement =
                anchor.EditAs?.Value
                ?? DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.TwoCell;
            if (placement == DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.Absolute) {
                return false;
            }
            if (placement != DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.OneCell
                || !ReferenceEquals(marker, anchor.ToMarker)) {
                return true;
            }

            return int.TryParse(anchor.FromMarker?.RowId?.Text, out int fromZeroBasedRow)
                && fromZeroBasedRow + 1 >= firstRow;
        }

        private static bool IsZeroOffsetDrawingBoundary(
            DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType marker) {
            return marker is DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker
                && long.TryParse(
                    marker.RowOffset?.Text,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out long rowOffset)
                && rowOffset == 0L;
        }

        private void ValidateChartFormulaCapacity(
            DocumentFormat.OpenXml.Packaging.DrawingsPart? drawingsPart,
            int firstRow,
            int count,
            bool rewriteUnqualifiedReferences) {
            if (drawingsPart == null) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Packaging.ChartPart chartPart in drawingsPart.ChartParts) {
                ValidateChartRootFormulaCapacity(
                    chartPart.ChartSpace,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences);
            }
            foreach (DocumentFormat.OpenXml.Packaging.ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                ValidateChartRootFormulaCapacity(
                    chartPart.ChartSpace,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences);
            }
        }

        private void ValidateDrawingShapeTextLinkCapacity(
            DocumentFormat.OpenXml.Packaging.DrawingsPart? drawingsPart,
            int firstRow,
            int count,
            bool rewriteUnqualifiedReferences) {
            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape
                in drawingsPart?.WorksheetDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>()
                ?? Enumerable.Empty<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>()) {
                ThrowIfFormulaReferenceOverflows(
                    shape.TextLink?.Value,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences);
            }
        }

        private void ValidateChartRootFormulaCapacity(
            OpenXmlPartRootElement? chartRoot,
            int firstRow,
            int count,
            bool rewriteUnqualifiedReferences) {
            if (chartRoot == null) {
                return;
            }

            foreach (OpenXmlLeafTextElement formula in chartRoot.Descendants<OpenXmlLeafTextElement>()
                .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                ThrowIfFormulaReferenceOverflows(
                    formula.Text,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences);
            }
        }

        private void ThrowIfFormulaReferenceOverflows(
            string? formula,
            int firstRow,
            int count,
            bool rewriteUnqualifiedReferences) {
            if (formula is not string formulaText || formulaText.Length == 0) {
                return;
            }

            bool overflow = false;
            RewriteFormulaReferencesOutsideStrings(formulaText, segment => {
                ReplaceFormulaRanges(segment, match => {
                    if (CanRewriteFormulaRangeQualifier(match, Name, rewriteUnqualifiedReferences)
                        && int.TryParse(match.Groups["startRow"].Value, out int first)
                        && int.TryParse(match.Groups["endRow"].Value, out int last)
                        && Math.Max(first, last) >= firstRow
                        && (long)Math.Max(first, last) + count > A1.MaxRows) {
                        overflow = true;
                    }
                    return match.Value;
                });
                ReplaceFormulaRowRanges(segment, match => {
                    if (CanRewriteFormulaRangeQualifier(match, Name, rewriteUnqualifiedReferences)
                        && int.TryParse(match.Groups["startRow"].Value, out int first)
                        && int.TryParse(match.Groups["endRow"].Value, out int last)
                        && Math.Max(first, last) >= firstRow
                        && (long)Math.Max(first, last) + count > A1.MaxRows) {
                        overflow = true;
                    }
                    return match.Value;
                });
                ReplaceFormulaReferences(segment, match => {
                    if (CanRewriteFormulaReference(
                            match,
                            Name,
                            allowAbsoluteRows: true,
                            allowOtherSheets: false,
                            out int row,
                            rewriteUnqualifiedReferences)
                        && row >= firstRow
                        && (long)row + count > A1.MaxRows) {
                        overflow = true;
                    }
                    return match.Value;
                });
                return segment;
            });

            if (overflow) {
                throw new InvalidOperationException("Inserting rows would move a formula reference beyond Excel's row limit.");
            }
        }

        private static void ValidateReferenceAttributesDoNotOverflow(OpenXmlElement root, int firstRow, int count) {
            foreach (OpenXmlElement element in root.Descendants().Prepend(root)) {
                if (element is SheetDimension || element is DataReference) {
                    continue;
                }
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (!string.Equals(attribute.LocalName, "ref", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(attribute.LocalName, "sqref", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    ValidateReferenceListDoesNotOverflow(attribute.Value, firstRow, count);
                }
            }
        }

        private static void ValidateReferenceListDoesNotOverflow(string? referenceList, int firstRow, int count) {
            if (referenceList is not string references || string.IsNullOrWhiteSpace(references)) {
                return;
            }

            foreach (ReferenceListPart part in SplitReferenceList(references)) {
                if (TryParseReference(part, out var bounds)
                    && bounds.r2 >= firstRow
                    && (long)bounds.r2 + count > A1.MaxRows) {
                    throw new InvalidOperationException(
                        $"Inserting rows would move reference '{part}' beyond Excel's row limit.");
                }
            }
        }

        private void InvalidateStructuralFormulaResults() {
            foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                Worksheet? worksheet = worksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }

                bool changed = false;
                foreach (CellFormula formula in worksheet.Descendants<CellFormula>()) {
                    if (formula.CalculateCell?.Value != true) {
                        formula.CalculateCell = true;
                        changed = true;
                    }
                }

                if (changed) {
                    worksheet.Save();
                }
            }

            _excelDocument.CleanupCalculationArtifacts(
                save: true,
                policy: ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
        }

        private void ShiftMergeCellsRows(int firstAffectedRow, int delta, int? lastDeletedRow = null) {
            MergeCells? merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null || delta == 0) {
                return;
            }

            uint count = 0;
            foreach (MergeCell merge in merges.Elements<MergeCell>().ToList()) {
                if (merge.Reference?.Value is not string reference
                    || !TryParseReference(reference, out var bounds)) {
                    count++;
                    continue;
                }

                if (!TryRemapShiftedReferenceRows(bounds, firstAffectedRow, delta, lastDeletedRow, out var remappedBounds)) {
                    count++;
                    continue;
                }

                if (remappedBounds == null) {
                    merge.Remove();
                    continue;
                }

                merge.Reference = ToReference(
                    remappedBounds.Value.r1,
                    remappedBounds.Value.c1,
                    remappedBounds.Value.r2,
                    remappedBounds.Value.c2);
                count++;
            }

            if (count == 0U) {
                merges.Remove();
            } else {
                merges.Count = count;
            }
        }

        private void ResetStructuralMutationCaches() {
            _excelDocument.ResetStructuralMutationCaches(_worksheetPart);
        }

        internal void ResetStructuralMutationCachesLocal() {
            _hasWorksheetMutations = true;
            _lastAccessedRow = null;
            _lastAccessedRowIndex = 0;
            _lastAccessedCell = null;
            _lastAccessedCellRowIndex = 0;
            _lastAccessedCellColumnIndex = 0;
            ClearHeaderCache();
            ClearFindFirstCache();
        }
    }
}
