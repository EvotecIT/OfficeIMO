using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static DirectColumnWritePlan[] CreateColumnWritePlans(DirectDataSetWorkbookModel model, DirectStylePlan stylePlan, CancellationToken ct) {
                var plans = new DirectColumnWritePlan[model.Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < plans.Length; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DirectDataSetSheetModel sheet = model.Sheets[i];
                    plans[i] = CreateColumnWritePlan(sheet.Table, sheet.UseCellValueNumberFormats, sheet.ColumnNumberFormats, stylePlan);
                }

                return plans;
            }

            private static void WriteDirectValueRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>>? overlayCellsByRow = null) {
                if (overlayCellsByRow != null) {
                    WriteDirectValueRowsWithOverlayCells(writer, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct, overlayCellsByRow);
                    return;
                }

                if (sheet.Table.TryGetExactDictionaryRows(out var exactDictionaryRows)) {
                    WriteExactDictionaryValueRows(writer, exactDictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetDictionaryRows(out var dictionaryRows)) {
                    WriteDictionaryValueRows(writer, dictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetLegacyDictionaryRows(out _)) {
                    WriteLegacyDictionaryValueRows(writer, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetCellValueRows(out DirectCellValueRows cellValueRows)) {
                    WriteDirectCellValueRows(writer, sheet, cellValueRows, rowCount, columnCount, startRowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectCellValueRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                DirectCellValueRows rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                object?[] values = rows.Values;

                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    if (columnCount == 4 && styleAttributes == null) {
                        WriteDirectCellValueRowsFourColumns(writer, rows, rowCount, startRowIndex, cellReferencePrefixes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                        return;
                    }

                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        int rowOffset = rows.GetRowOffset(sourceRowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            WriteCell(writer, rowReference, cellReferencePrefixes[c], values[rowOffset + c], styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        int rowOffset = rows.GetRowOffset(sourceRowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = values[rowOffset + c];
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    int rowOffset = rows.GetRowOffset(sourceRowIndex);
                    for (int c = 0; c < columnCount; c++) {
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], values[rowOffset + c], styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectCellValueRowsFourColumns(
                TextWriter writer,
                DirectCellValueRows rows,
                int rowCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                object?[] values = rows.Values;
                string prefix0 = cellReferencePrefixes[0];
                string prefix1 = cellReferencePrefixes[1];
                string prefix2 = cellReferencePrefixes[2];
                string prefix3 = cellReferencePrefixes[3];
                DirectCellValueKind kind0 = cellValueKinds[0];
                DirectCellValueKind kind1 = cellValueKinds[1];
                DirectCellValueKind kind2 = cellValueKinds[2];
                DirectCellValueKind kind3 = cellValueKinds[3];

                if (kind0 == DirectCellValueKind.Int32
                    && kind1 == DirectCellValueKind.String
                    && kind2 == DirectCellValueKind.String
                    && kind3 == DirectCellValueKind.Double) {
                    if (rows.ValuesMatchColumnTypes) {
                        WriteDirectCellValueRowsExactIntStringStringDouble(writer, values, rowCount, startRowIndex, prefix0, prefix1, prefix2, prefix3, sharedStrings, ct);
                    } else {
                        WriteDirectCellValueRowsIntStringStringDouble(writer, values, rowCount, startRowIndex, prefix0, prefix1, prefix2, prefix3, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    }

                    return;
                }

                for (int sourceRowIndex = 0, offset = 0; sourceRowIndex < rowCount; sourceRowIndex++, offset += 4) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    WriteCell(writer, rowReference, prefix0, values[offset], null, kind0, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    WriteCell(writer, rowReference, prefix1, values[offset + 1], null, kind1, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    WriteCell(writer, rowReference, prefix2, values[offset + 2], null, kind2, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    WriteCell(writer, rowReference, prefix3, values[offset + 3], null, kind3, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectCellValueRowsIntStringStringDouble(
                TextWriter writer,
                object?[] values,
                int rowCount,
                int startRowIndex,
                string prefix0,
                string prefix1,
                string prefix2,
                string prefix3,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                for (int sourceRowIndex = 0, offset = 0; sourceRowIndex < rowCount; sourceRowIndex++, offset += 4) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");

                    writer.Write(prefix0);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (values[offset] is int intValue) {
                        WriteRawValueCell(writer, intValue);
                    } else {
                        WriteCellValue(writer, values[offset], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix1);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (values[offset + 1] is string stringValue1) {
                        WriteStringCellValue(writer, stringValue1, sharedStrings);
                    } else {
                        WriteCellValue(writer, values[offset + 1], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix2);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (values[offset + 2] is string stringValue2) {
                        WriteStringCellValue(writer, stringValue2, sharedStrings);
                    } else {
                        WriteCellValue(writer, values[offset + 2], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix3);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (values[offset + 3] is double doubleValue) {
                        WriteRawValueCell(writer, doubleValue);
                    } else {
                        WriteCellValue(writer, values[offset + 3], DirectCellValueKind.Double, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectCellValueRowsExactIntStringStringDouble(
                TextWriter writer,
                object?[] values,
                int rowCount,
                int startRowIndex,
                string prefix0,
                string prefix1,
                string prefix2,
                string prefix3,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                for (int sourceRowIndex = 0, offset = 0; sourceRowIndex < rowCount; sourceRowIndex++, offset += 4) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");

                    writer.Write(prefix0);
                    writer.Write(rowReference);
                    writer.Write('"');
                    WriteRawValueCell(writer, (int)values[offset]!);

                    writer.Write(prefix1);
                    writer.Write(rowReference);
                    writer.Write('"');
                    WriteStringCellValue(writer, (string)values[offset + 1]!, sharedStrings);

                    writer.Write(prefix2);
                    writer.Write(rowReference);
                    writer.Write('"');
                    WriteStringCellValue(writer, (string)values[offset + 2]!, sharedStrings);

                    writer.Write(prefix3);
                    writer.Write(rowReference);
                    writer.Write('"');
                    WriteRawValueCell(writer, (double)values[offset + 3]!);

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectValueRowsWithOverlayCells(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>> overlayCellsByRow) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (overlayCellsByRow.ContainsKey(rowIndex)) {
                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                            }

                            WriteOverlayCellsForRow(writer, overlayCellsByRow, rowIndex, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                            rowStarted = true;
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    WriteOverlayCellsForRow(writer, overlayCellsByRow, rowIndex, stylePlan, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteExactDictionaryValueRows(
                TextWriter writer,
                IReadOnlyList<Dictionary<string, object?>> rows,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindExactDictionaryRows(writer, rows, rowCount, columnCount, rowIndex, cellReferencePrefixes, columnNames, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        Dictionary<string, object?> row = rows[sourceRowIndex];
                        for (int c = 0; c < columnCount; c++) {
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    Dictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindExactDictionaryRows(
                TextWriter writer,
                IReadOnlyList<Dictionary<string, object?>> rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
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
                        Dictionary<string, object?> row = rows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
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
                    Dictionary<string, object?> row = rows[sourceRowIndex];
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

                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDictionaryValueRows(
                TextWriter writer,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindDictionaryRows(writer, rows, rowCount, columnCount, rowIndex, cellReferencePrefixes, columnNames, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                        for (int c = 0; c < columnCount; c++) {
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindDictionaryRows(
                TextWriter writer,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
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
                        IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
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
                    IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
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

                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteLegacyDictionaryValueRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindLegacyDictionaryRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindLegacyDictionaryRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
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
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
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

                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }
        }

    }
}
