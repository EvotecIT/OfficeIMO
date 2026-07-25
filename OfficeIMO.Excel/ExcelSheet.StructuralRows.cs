using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

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
        /// would split an array formula or PivotTable output range.
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
        /// Thrown when the deletion would remove an owned table boundary or intersect an array-formula or PivotTable
        /// output range that cannot be rebuilt safely.
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

            ValidateStructuralRowControlSafety();
            ValidateRowInsertionAgainstArrayFormulas(firstRow);
            ValidateRowInsertionAgainstPivotOutputs(firstRow);
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            uint maxShiftedRow = sheetData?.Elements<Row>()
                .Where(item => item.RowIndex?.Value >= (uint)firstRow)
                .Select(item => item.RowIndex?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() ?? 0U;
            if (maxShiftedRow > 0U && (long)maxShiftedRow + count > A1.MaxRows) {
                throw new InvalidOperationException("Inserting rows would move worksheet content beyond Excel's row limit.");
            }
            ValidateStructuralRowReferenceCapacity(firstRow, count);
            MaterializeWorkbookSharedFormulasForStructuralEdit();

            if (sheetData != null) {
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
            ValidateStructuralRowControlSafety();
            ValidateRowDeletionAgainstOwnedRanges(firstRow, lastRemovedRow);
            ValidateWorkbookSharedFormulasForStructuralEdit();
            MaterializeWorkbookSharedFormulasForStructuralEdit();
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (Row row in sheetData.Elements<Row>().ToList()) {
                    if (row.RowIndex == null) {
                        continue;
                    }

                    int rowIndex = checked((int)row.RowIndex.Value);
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
                .Where(item => item.FormulaType?.Value == CellFormulaValues.Array)) {
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
                    throw new InvalidOperationException(
                        $"Cannot insert rows through array formula range '{reference}'. Insert before or after the complete array range.");
                }
            }
        }

        private void ValidateRowDeletionAgainstOwnedRanges(int firstDeletedRow, int lastDeletedRow) {
            foreach (CellFormula formula in WorksheetRoot.Descendants<CellFormula>()
                .Where(item => item.FormulaType?.Value == CellFormulaValues.Array)) {
                if (formula.Reference?.Value is not string reference
                    || !A1.TryParseRange(
                        reference.Replace("$", string.Empty),
                        out int arrayFirstRow,
                        out _,
                        out int arrayLastRow,
                        out _)) {
                    continue;
                }

                bool deletesOwner = arrayFirstRow >= firstDeletedRow && arrayFirstRow <= lastDeletedRow;
                bool deletesWholeArray = firstDeletedRow <= arrayFirstRow && lastDeletedRow >= arrayLastRow;
                if (deletesOwner && !deletesWholeArray) {
                    throw new InvalidOperationException(
                        $"Cannot delete the owner row of array formula range '{reference}' while part of the array survives.");
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

            IEnumerable<DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart> pivotCaches =
                WorkbookPartRoot.WorksheetParts
                    .SelectMany(part => part.PivotTableParts)
                    .Select(part => part.PivotTableCacheDefinitionPart)
                    .Where(part => part != null)
                    .Cast<DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart>()
                    .Distinct();
            foreach (var cachePart in pivotCaches) {
                WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (source?.Reference?.Value is not string reference
                    || !string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    || !A1.TryParseRange(reference.Replace("$", string.Empty), out int sourceFirstRow, out _, out int sourceLastRow, out _)) {
                    continue;
                }

                if (firstDeletedRow <= sourceFirstRow && lastDeletedRow >= sourceLastRow) {
                    throw new InvalidOperationException(
                        $"Cannot delete the complete source range '{reference}' of a pivot cache. Update or remove the pivot source first.");
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

                foreach (Hyperlink hyperlink in worksheet.Descendants<Hyperlink>()) {
                    if (string.IsNullOrWhiteSpace(hyperlink.Id?.Value)) {
                        ThrowIfFormulaReferenceOverflows(hyperlink.Location?.Value, firstRow, count, rewriteUnqualified);
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

                ValidateChartFormulaCapacity(worksheetPart.DrawingsPart, firstRow, count);
            }

            ValidatePivotReferenceCapacity(firstRow, count);

            foreach (DocumentFormat.OpenXml.Packaging.ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                ValidateChartFormulaCapacity(chartsheetPart.DrawingsPart, firstRow, count);
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

            foreach (Break pageBreak in WorksheetRoot.GetFirstChild<RowBreaks>()?.Elements<Break>() ?? Enumerable.Empty<Break>()) {
                if (pageBreak.Id?.Value is uint row && row >= firstRow && (long)row + count > A1.MaxRows) {
                    throw new InvalidOperationException("Inserting rows would move a page break beyond Excel's row limit.");
                }
            }

            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType marker in
                _worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType>()
                ?? Enumerable.Empty<DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType>()) {
                if (int.TryParse(marker.RowId?.Text, out int zeroBasedRow)
                    && zeroBasedRow + 1 >= firstRow
                    && (long)zeroBasedRow + 1L + count > A1.MaxRows) {
                    throw new InvalidOperationException("Inserting rows would move a drawing anchor beyond Excel's row limit.");
                }
            }
        }

        private void ValidatePivotReferenceCapacity(int firstRow, int count) {
            IEnumerable<DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart> cacheParts =
                WorkbookPartRoot.WorksheetParts
                    .SelectMany(worksheetPart => worksheetPart.PivotTableParts)
                    .Select(pivotPart => pivotPart.PivotTableCacheDefinitionPart)
                    .Where(cachePart => cachePart != null)
                    .Cast<DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart>()
                    .Distinct();
            foreach (DocumentFormat.OpenXml.Packaging.PivotTableCacheDefinitionPart cachePart in cacheParts) {
                WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (source != null
                    && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)) {
                    ValidateReferenceListDoesNotOverflow(source.Reference?.Value, firstRow, count);
                }
            }

            foreach (DocumentFormat.OpenXml.Packaging.PivotTablePart pivotPart in _worksheetPart.PivotTableParts) {
                ValidateReferenceListDoesNotOverflow(
                    pivotPart.PivotTableDefinition?.Location?.Reference?.Value,
                    firstRow,
                    count);
            }
        }

        private void ValidateStructuralRowControlSafety() {
            if (WorksheetRoot.Descendants<Controls>().Any()
                || _worksheetPart.ControlPropertiesParts.Any()) {
                throw new InvalidOperationException(
                    "Cannot edit rows on a worksheet containing form controls because their anchors and linked cells cannot yet be remapped safely.");
            }
        }

        private void ValidateChartFormulaCapacity(
            DocumentFormat.OpenXml.Packaging.DrawingsPart? drawingsPart,
            int firstRow,
            int count) {
            if (drawingsPart == null) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Packaging.ChartPart chartPart in drawingsPart.ChartParts) {
                ValidateChartRootFormulaCapacity(chartPart.ChartSpace, firstRow, count);
            }
            foreach (DocumentFormat.OpenXml.Packaging.ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                ValidateChartRootFormulaCapacity(chartPart.ChartSpace, firstRow, count);
            }
        }

        private void ValidateChartRootFormulaCapacity(
            OpenXmlPartRootElement? chartRoot,
            int firstRow,
            int count) {
            if (chartRoot == null) {
                return;
            }

            foreach (OpenXmlLeafTextElement formula in chartRoot.Descendants<OpenXmlLeafTextElement>()
                .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                ThrowIfFormulaReferenceOverflows(
                    formula.Text,
                    firstRow,
                    count,
                    rewriteUnqualifiedReferences: false);
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
                if (element is SheetDimension) {
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

            merges.Count = count;
        }

        private void ResetStructuralMutationCaches() {
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
