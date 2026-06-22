using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteWorksheet(ZipArchive archive, DirectDataSetSheetModel sheet, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings, DirectStylePlan stylePlan, DirectColumnWritePlan columnWritePlan, CancellationToken ct, string? worksheetPath = null, string? tableRelationshipId = null) {
                var entry = archive.CreateEntry(worksheetPath ?? "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                var metadata = sheet.Metadata;
                if (!string.IsNullOrEmpty(metadata?.SheetPropertiesXml)) {
                    writer.Write(metadata!.SheetPropertiesXml);
                }

                writer.Write("<dimension ref=\"");
                WriteEscaped(writer, GetWorksheetDimension(sheet));
                writer.Write("\"/>");
                if (!string.IsNullOrEmpty(metadata?.SheetViewsXml)) {
                    writer.Write(metadata!.SheetViewsXml);
                }

                if (!string.IsNullOrEmpty(metadata?.SheetFormatPropertiesXml)) {
                    writer.Write(metadata!.SheetFormatPropertiesXml);
                }

                WriteColumns(writer, sheet.ColumnWidths);
                writer.Write("<sheetData>");
                int columnCount = sheet.Table.ColumnCount;
                int rowCount = sheet.Table.RowCount;
                string[] cellReferencePrefixes = CreateCellReferencePrefixes(columnCount);
                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                DirectCellValueKind[] cellValueKinds = columnWritePlan.CellValueKinds;
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = 1;
                var overlayCells = metadata?.OverlayCells;
                var overlayCellsByRow = CreateOverlayCellsByRow(overlayCells);
                int directLastRow = rowCount + (sheet.IncludeHeaders ? 1 : 0);
                bool hasInlineOverlayCells = HasOverlayCellsAtOrBefore(overlayCells, directLastRow);
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int c = 0; c < columnCount; c++) {
                        WriteCell(writer, headerRowReference, cellReferencePrefixes[c], sheet.Table.GetColumnName(c), null, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    WriteOverlayCellsForRow(writer, overlayCellsByRow, 1, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    writer.Write("</row>");
                    rowIndex++;
                }

                bool hasBufferedRows = sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows);
                bool hasCellValueRows = sheet.Table.TryGetCellValueRows(out DirectCellValueRows cellValueRows);
                if (hasInlineOverlayCells) {
                    WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct, overlayCellsByRow);
                    rowIndex += rowCount;
                } else if (hasBufferedRows
                    && !sheet.OmitBlankCells
                    && styleAttributes == null
                    && valueStyleColumns == null
                    && AllColumnsUseCellValueKind(cellValueKinds, DirectCellValueKind.String)) {
                    WritePlainStringBufferedRows(writer, bufferedRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, sharedStrings, ct);
                    rowIndex += rowCount;
                } else if (hasBufferedRows
                    && !sheet.OmitBlankCells
                    && valueStyleColumns == null) {
                    WriteFixedKindBufferedRows(writer, bufferedRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    rowIndex += rowCount;
                } else if (hasCellValueRows
                    && !sheet.OmitBlankCells
                    && valueStyleColumns == null) {
                    WriteFixedKindCellValueRows(writer, cellValueRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    rowIndex += rowCount;
                } else if (valueStyleColumns == null) {
                    if (hasBufferedRows) {
                        if (sheet.OmitBlankCells) {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                bool rowStarted = false;
                                string rowReference = InvariantNumberText.Get(rowIndex);
                                object?[] bufferedRow = bufferedRows[sourceRowIndex];
                                for (int c = 0; c < columnCount; c++) {
                                    object? value = bufferedRow[c];
                                    if (IsBlankCellValue(value)) {
                                        continue;
                                    }

                                    if (!rowStarted) {
                                        writer.Write("<row r=\"");
                                        writer.Write(rowReference);
                                        writer.Write("\">");
                                        rowStarted = true;
                                    }

                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                                }

                                if (rowStarted) {
                                    writer.Write("</row>");
                                }

                                rowIndex++;
                            }
                        } else {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                string rowReference = InvariantNumberText.Get(rowIndex);
                                object?[] bufferedRow = bufferedRows[sourceRowIndex];
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                for (int c = 0; c < columnCount; c++) {
                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], bufferedRow[c], styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                                }

                                writer.Write("</row>");
                                rowIndex++;
                            }
                        }
                    } else if (sheet.Table.HasSourceRows) {
                        if (sheet.OmitBlankCells) {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                bool rowStarted = false;
                                string rowReference = InvariantNumberText.Get(rowIndex);
                                DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                                for (int c = 0; c < columnCount; c++) {
                                    object? value = sourceRow[c];
                                    if (IsBlankCellValue(value)) {
                                        continue;
                                    }

                                    if (!rowStarted) {
                                        writer.Write("<row r=\"");
                                        writer.Write(rowReference);
                                        writer.Write("\">");
                                        rowStarted = true;
                                    }

                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                                }

                                if (rowStarted) {
                                    writer.Write("</row>");
                                }

                                rowIndex++;
                            }
                        } else {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                string rowReference = InvariantNumberText.Get(rowIndex);
                                DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                for (int c = 0; c < columnCount; c++) {
                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], sourceRow[c], styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                                }

                                writer.Write("</row>");
                                rowIndex++;
                            }
                        }
                    } else {
                        WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns: null, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                        rowIndex += rowCount;
                    }
                } else if (hasBufferedRows) {
                    if (sheet.OmitBlankCells) {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            bool rowStarted = false;
                            string rowReference = InvariantNumberText.Get(rowIndex);
                            object?[] bufferedRow = bufferedRows[sourceRowIndex];
                            for (int c = 0; c < columnCount; c++) {
                                object? value = bufferedRow[c];
                                if (IsBlankCellValue(value)) {
                                    continue;
                                }

                                if (!rowStarted) {
                                    writer.Write("<row r=\"");
                                    writer.Write(rowReference);
                                    writer.Write("\">");
                                    rowStarted = true;
                                }

                                WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                            }

                            if (rowStarted) {
                                writer.Write("</row>");
                            }

                            rowIndex++;
                        }
                    } else {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            string rowReference = InvariantNumberText.Get(rowIndex);
                            object?[] bufferedRow = bufferedRows[sourceRowIndex];
                            writer.Write("<row r=\"");
                            writer.Write(rowReference);
                            writer.Write("\">");
                            for (int c = 0; c < columnCount; c++) {
                                WriteCell(writer, rowReference, cellReferencePrefixes[c], bufferedRow[c], styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                            }

                            writer.Write("</row>");
                            rowIndex++;
                        }
                    }
                } else if (sheet.Table.HasSourceRows) {
                    if (sheet.OmitBlankCells) {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            bool rowStarted = false;
                            string rowReference = InvariantNumberText.Get(rowIndex);
                            DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                            for (int c = 0; c < columnCount; c++) {
                                object? value = sourceRow[c];
                                if (IsBlankCellValue(value)) {
                                    continue;
                                }

                                if (!rowStarted) {
                                    writer.Write("<row r=\"");
                                    writer.Write(rowReference);
                                    writer.Write("\">");
                                    rowStarted = true;
                                }

                                WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                            }

                            if (rowStarted) {
                                writer.Write("</row>");
                            }

                            rowIndex++;
                        }
                    } else {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            string rowReference = InvariantNumberText.Get(rowIndex);
                            DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                            writer.Write("<row r=\"");
                            writer.Write(rowReference);
                            writer.Write("\">");
                            for (int c = 0; c < columnCount; c++) {
                                WriteCell(writer, rowReference, cellReferencePrefixes[c], sourceRow[c], styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                            }

                            writer.Write("</row>");
                            rowIndex++;
                        }
                    }
                } else {
                    WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    rowIndex += rowCount;
                }

                WriteOverlayCells(writer, overlayCells, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, directLastRow);
                writer.Write("</sheetData>");
                if (!string.IsNullOrEmpty(metadata?.SheetProtectionXml)) {
                    writer.Write(metadata!.SheetProtectionXml);
                }

                if (!string.IsNullOrEmpty(metadata?.AutoFilterXml)) {
                    writer.Write(metadata!.AutoFilterXml);
                }

                if (metadata != null) {
                    for (int i = 0; i < metadata.ConditionalFormattingXml.Count; i++) {
                        writer.Write(metadata.ConditionalFormattingXml[i]);
                    }
                }

                if (!string.IsNullOrEmpty(metadata?.DataValidationsXml)) {
                    writer.Write(metadata!.DataValidationsXml);
                }

                if (metadata != null) {
                    for (int i = 0; i < metadata.PostDataValidationXml.Count; i++) {
                        writer.Write(metadata.PostDataValidationXml[i]);
                    }
                }

                if (!string.IsNullOrEmpty(metadata?.DrawingXml)) {
                    writer.Write(metadata!.DrawingXml);
                }

                if (sheet.HasTable) {
                    writer.Write("<tableParts count=\"1\"><tablePart r:id=\"");
                    WriteEscaped(writer, tableRelationshipId ?? "rId1");
                    writer.Write("\"/></tableParts>");
                }

                writer.Write("</worksheet>");
            }

            private static void WriteFixedKindCellValueRows(
                TextWriter writer,
                DirectCellValueRows cellValueRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    object?[] values = cellValueRows.Values;
                    int rowOffset = 0;
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            WriteCellValue(writer, values[rowOffset + c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                        rowOffset += columnCount;
                    }

                    return;
                }

                object?[] styledValues = cellValueRows.Values;
                int styledRowOffset = 0;
                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        WriteCellValue(writer, styledValues[styledRowOffset + c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                    styledRowOffset += columnCount;
                }
            }

            private static void WriteFixedKindBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        object?[] bufferedRow = bufferedRows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            WriteCellValue(writer, bufferedRow[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] bufferedRow = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        WriteCellValue(writer, bufferedRow[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WritePlainStringBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                if (sharedStrings == null) {
                    WritePlainInlineStringBufferedRows(writer, bufferedRows, rowCount, columnCount, startRowIndex, cellReferencePrefixes, ct);
                    return;
                }

                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] bufferedRow = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        object? value = bufferedRow[c];
                        if (value is string text) {
                            WriteStringCellValue(writer, text, sharedStrings);
                        } else if (value == null || value == DBNull.Value) {
                            writer.Write(" t=\"str\"><v/></c>");
                        } else {
                            WriteStringCell(writer, value.ToString() ?? string.Empty, validateLength: true);
                        }
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WritePlainInlineStringBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] bufferedRow = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        object? value = bufferedRow[c];
                        if (value is string text) {
                            WriteStringCell(writer, text, validateLength: true);
                        } else if (value == null || value == DBNull.Value) {
                            writer.Write(" t=\"str\"><v/></c>");
                        } else {
                            WriteStringCell(writer, value.ToString() ?? string.Empty, validateLength: true);
                        }
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static bool AllColumnsUseCellValueKind(DirectCellValueKind[] cellValueKinds, DirectCellValueKind expected) {
                for (int i = 0; i < cellValueKinds.Length; i++) {
                    if (cellValueKinds[i] != expected) {
                        return false;
                    }
                }

                return true;
            }

            private static bool IsBlankCellValue(object? value) {
                return value == null || value == DBNull.Value;
            }

            private static string GetWorksheetDimension(DirectDataSetSheetModel sheet) {
                var overlayCells = sheet.Metadata?.OverlayCells;
                if (overlayCells == null || overlayCells.Count == 0) {
                    return sheet.Range.Length == 0 ? "A1" : sheet.Range;
                }

                int lastRow = sheet.Table.RowCount + (sheet.IncludeHeaders ? 1 : 0);
                int lastColumn = sheet.Table.ColumnCount;
                for (int i = 0; i < overlayCells.Count; i++) {
                    if (overlayCells[i].IsDeleted) {
                        continue;
                    }

                    lastRow = Math.Max(lastRow, overlayCells[i].Row);
                    lastColumn = Math.Max(lastColumn, overlayCells[i].Column);
                }

                return "A1:" + A1.CellReference(lastRow, lastColumn);
            }

            private static void WriteOverlayCells(
                TextWriter writer,
                IReadOnlyList<DirectOverlayCell>? overlayCells,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                int skipRowsAtOrBefore = 0) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return;
                }

                int row = -1;
                for (int i = 0; i < overlayCells.Count; i++) {
                    var cell = overlayCells[i];
                    if (cell.IsDeleted || cell.Row <= skipRowsAtOrBefore) {
                        continue;
                    }

                    if (cell.Row != row) {
                        if (row != -1) {
                            writer.Write("</row>");
                        }

                        row = cell.Row;
                        writer.Write("<row r=\"");
                        writer.Write(InvariantNumberText.Get(row));
                        writer.Write("\">");
                    }

                    WriteCell(
                        writer,
                        InvariantNumberText.Get(cell.Row),
                        "<c r=\"" + A1.ColumnIndexToLetters(cell.Column),
                        cell.Value,
                        CreateOverlayStyleAttribute(cell, stylePlan),
                        dateTimeOffsetWriteStrategy,
                        dateSystem,
                        sharedStrings);
                }

                if (row != -1) {
                    writer.Write("</row>");
                }
            }

            private static Dictionary<int, IReadOnlyList<DirectOverlayCell>>? CreateOverlayCellsByRow(IReadOnlyList<DirectOverlayCell>? overlayCells) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return null;
                }

                return overlayCells
                    .Where(static cell => !cell.IsDeleted)
                    .GroupBy(static cell => cell.Row)
                    .ToDictionary(
                        static group => group.Key,
                        static group => (IReadOnlyList<DirectOverlayCell>)group.OrderBy(static cell => cell.Column).ToArray());
            }

            private static bool HasOverlayCellsAtOrBefore(IReadOnlyList<DirectOverlayCell>? overlayCells, int row) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return false;
                }

                for (int i = 0; i < overlayCells.Count; i++) {
                    if (!overlayCells[i].IsDeleted && overlayCells[i].Row <= row) {
                        return true;
                    }
                }

                return false;
            }

            private static bool WriteOverlayCellsForRow(
                TextWriter writer,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>>? overlayCellsByRow,
                int row,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings) {
                if (overlayCellsByRow == null || !overlayCellsByRow.TryGetValue(row, out var rowOverlayCells)) {
                    return false;
                }

                string rowReference = InvariantNumberText.Get(row);
                for (int i = 0; i < rowOverlayCells.Count; i++) {
                    var cell = rowOverlayCells[i];
                    if (cell.IsDeleted) {
                        continue;
                    }

                    WriteCell(
                        writer,
                        rowReference,
                        "<c r=\"" + A1.ColumnIndexToLetters(cell.Column),
                        cell.Value,
                        CreateOverlayStyleAttribute(cell, stylePlan),
                        dateTimeOffsetWriteStrategy,
                        dateSystem,
                        sharedStrings);
                }

                return true;
            }

            private static string? CreateOverlayStyleAttribute(DirectOverlayCell cell, DirectStylePlan stylePlan) {
                if (!string.IsNullOrWhiteSpace(cell.NumberFormat)) {
                    return stylePlan.GetStyleAttribute(cell.NumberFormat!);
                }

                return cell.StyleIndex == 0U ? " s=\"0\"" : null;
            }

            private static string[] CreateCellReferencePrefixes(int columnCount) {
                if ((uint)columnCount <= CachedCellReferencePrefixColumnLimit) {
                    return CellReferencePrefixCache[columnCount];
                }

                var columns = new string[columnCount];
                for (int i = 0; i < columnCount; i++) {
                    columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                }

                return columns;
            }

            private static string[][] CreateCellReferencePrefixCache() {
                var cache = new string[CachedCellReferencePrefixColumnLimit + 1][];
                cache[0] = Array.Empty<string>();
                for (int columnCount = 1; columnCount < cache.Length; columnCount++) {
                    var columns = new string[columnCount];
                    for (int i = 0; i < columnCount; i++) {
                        columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                    }

                    cache[columnCount] = columns;
                }

                return cache;
            }

            private static DirectColumnWritePlan CreateColumnWritePlan(
                DirectDataSetTableModel table,
                bool useCellValueNumberFormats,
                IReadOnlyList<string?>? columnNumberFormats,
                DirectStylePlan stylePlan) {
                int columnCount = table.ColumnCount;
                var kinds = new DirectCellValueKind[columnCount];
                string?[]? styleAttributes = null;
                bool[]? valueStyleColumns = null;

                for (int i = 0; i < columnCount; i++) {
                    Type dataType = table.GetColumnType(i);
                    DirectCellValueKind cellValueKind = GetCellValueKind(dataType);
                    bool useValueStyle = dataType == typeof(object);
                    if (useValueStyle && TryInferObjectColumnCellValueKind(table, i, out DirectCellValueKind inferredKind)) {
                        cellValueKind = inferredKind;
                        useValueStyle = false;
                    }

                    kinds[i] = cellValueKind;

                    string? styleAttribute = null;
                    if (columnNumberFormats != null
                        && i < columnNumberFormats.Count
                        && columnNumberFormats[i] is string numberFormat
                        && !string.IsNullOrWhiteSpace(numberFormat)) {
                        styleAttribute = stylePlan.GetStyleAttribute(numberFormat);
                    }

                    styleAttribute ??= GetStyleAttribute(cellValueKind, useCellValueNumberFormats);
                    if (styleAttribute != null) {
                        styleAttributes ??= new string?[columnCount];
                        styleAttributes[i] = styleAttribute;
                    }

                    if (useValueStyle) {
                        valueStyleColumns ??= new bool[columnCount];
                        valueStyleColumns[i] = true;
                    }
                }

                return new DirectColumnWritePlan(kinds, styleAttributes, valueStyleColumns);
            }

            private static bool TryInferObjectColumnCellValueKind(DirectDataSetTableModel table, int columnIndex, out DirectCellValueKind inferredKind) {
                inferredKind = DirectCellValueKind.Object;
                int rowCount = table.RowCount;
                bool sawValue = false;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    object? value = table.GetValue(rowIndex, columnIndex);
                    if (value == null || value == DBNull.Value) {
                        continue;
                    }

                    DirectCellValueKind valueKind = GetCellValueKind(value.GetType());
                    if (valueKind == DirectCellValueKind.Object) {
                        return false;
                    }

                    if (!sawValue) {
                        inferredKind = valueKind;
                        sawValue = true;
                        continue;
                    }

                    if (inferredKind != valueKind) {
                        return false;
                    }
                }

                return sawValue;
            }

            private static DirectCellValueKind GetCellValueKind(Type dataType) {
                if (dataType == typeof(DirectFormulaCellValue)) return DirectCellValueKind.Formula;
                if (dataType == typeof(string)) return DirectCellValueKind.String;
                if (dataType == typeof(bool)) return DirectCellValueKind.Boolean;
                if (dataType == typeof(DateTime)) return DirectCellValueKind.DateTime;
                if (dataType == typeof(DateTimeOffset)) return DirectCellValueKind.DateTimeOffset;
                if (dataType == typeof(TimeSpan)) return DirectCellValueKind.TimeSpan;
                if (dataType == typeof(double)) return DirectCellValueKind.Double;
                if (dataType == typeof(float)) return DirectCellValueKind.Float;
                if (dataType == typeof(decimal)) return DirectCellValueKind.Decimal;
                if (dataType == typeof(sbyte)) return DirectCellValueKind.SByte;
                if (dataType == typeof(byte)) return DirectCellValueKind.Byte;
                if (dataType == typeof(short)) return DirectCellValueKind.Int16;
                if (dataType == typeof(ushort)) return DirectCellValueKind.UInt16;
                if (dataType == typeof(int)) return DirectCellValueKind.Int32;
                if (dataType == typeof(uint)) return DirectCellValueKind.UInt32;
                if (dataType == typeof(long)) return DirectCellValueKind.Int64;
                if (dataType == typeof(ulong)) return DirectCellValueKind.UInt64;
#if NET6_0_OR_GREATER
                if (dataType == typeof(DateOnly)) return DirectCellValueKind.DateOnly;
                if (dataType == typeof(TimeOnly)) return DirectCellValueKind.TimeOnly;
#endif
                return DirectCellValueKind.Object;
            }

            private static void WriteColumns(TextWriter writer, double[]? columnWidths) {
                if (columnWidths == null || columnWidths.Length == 0) {
                    return;
                }

                bool started = false;
                for (int i = 0; i < columnWidths.Length; i++) {
                    double width = NormalizeColumnWidth(columnWidths[i]);
                    if (width <= 0D) {
                        continue;
                    }

                    if (!started) {
                        writer.Write("<cols>");
                        started = true;
                    }

                    writer.Write("<col min=\"");
                    WriteInvariant(writer, i + 1);
                    writer.Write("\" max=\"");
                    WriteInvariant(writer, i + 1);
                    writer.Write("\" width=\"");
                    WriteColumnWidth(writer, width);
                    writer.Write("\" bestFit=\"1\" customWidth=\"1\"/>");
                }

                if (started) {
                    writer.Write("</cols>");
                }
            }

            private static double NormalizeColumnWidth(double width) {
                if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) {
                    return 0D;
                }

                return Math.Min(width, 255D);
            }
        }

    }
}
