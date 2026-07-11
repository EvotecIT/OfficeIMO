using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteCompactDirectValueRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                if (valueStyleColumns == null) {
                    if (sheet.Table.TryGetCellValueRows(out DirectCellValueRows cellValueRows)) {
                        WriteCompactFixedKindCellValueRows(
                            writer,
                            cellValueRows,
                            rowCount,
                            columnCount,
                            styleAttributes,
                            cellValueKinds,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings,
                            ct);
                        return;
                    }

                    if (sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows)) {
                        WriteCompactFixedKindBufferedRows(
                            writer,
                            bufferedRows,
                            rowCount,
                            columnCount,
                            styleAttributes,
                            cellValueKinds,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings,
                            ct);
                        return;
                    }
                }

                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    writer.Write("<row>");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        object? value = sheet.Table.GetValue(rowIndex, columnIndex);
                        writer.Write("<c");
                        bool useValueStyle = valueStyleColumns?[columnIndex] ?? false;
                        string? styleAttribute = styleAttributes?[columnIndex]
                            ?? (useValueStyle ? CreateStyleAttributeForValue(value, sheet.UseCellValueNumberFormats) : null);
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        if (useValueStyle) {
                            WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        } else {
                            WriteCellValue(writer, value, cellValueKinds[columnIndex], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                        }
                    }

                    writer.Write("</row>");
                }
            }

            private static void WriteCompactFixedKindCellValueRows(
                TextWriter writer,
                DirectCellValueRows cellValueRows,
                int rowCount,
                int columnCount,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                object?[] values = cellValueRows.Values;
                bool canCancel = ct.CanBeCanceled;
                if (columnCount == 8 && IsIntStringStringDateTimeDoubleIntBooleanStringPlan(cellValueKinds)) {
                    for (int rowIndex = 0, offset = 0; rowIndex < rowCount; rowIndex++, offset += 8) {
                        if (canCancel) ct.ThrowIfCancellationRequested();
                        WriteCompactIntStringStringDateTimeDoubleIntBooleanStringRow(
                            writer,
                            values,
                            offset,
                            styleAttributes,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings);
                    }

                    return;
                }

                int rowOffset = 0;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++, rowOffset += columnCount) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    writer.Write("<row>");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        writer.Write("<c");
                        string? styleAttribute = styleAttributes?[columnIndex];
                        if (styleAttribute != null) writer.Write(styleAttribute);
                        WriteCellValue(writer, values[rowOffset + columnIndex], cellValueKinds[columnIndex], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                }
            }

            private static void WriteCompactFixedKindBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                if (columnCount == 8 && IsIntStringStringDateTimeDoubleIntBooleanStringPlan(cellValueKinds)) {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) ct.ThrowIfCancellationRequested();
                        WriteCompactIntStringStringDateTimeDoubleIntBooleanStringRow(
                            writer,
                            bufferedRows[rowIndex],
                            0,
                            styleAttributes,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings);
                    }

                    return;
                }

                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    object?[] values = bufferedRows[rowIndex];
                    writer.Write("<row>");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        writer.Write("<c");
                        string? styleAttribute = styleAttributes?[columnIndex];
                        if (styleAttribute != null) writer.Write(styleAttribute);
                        WriteCellValue(writer, values[columnIndex], cellValueKinds[columnIndex], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                }
            }

            private static void WriteCompactIntStringStringDateTimeDoubleIntBooleanStringRow(
                TextWriter writer,
                object?[] values,
                int offset,
                string?[]? styleAttributes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings) {
                writer.Write("<row><c");
                if (styleAttributes?[0] is string style0) writer.Write(style0);
                if (values[offset] is int intValue0) WriteRawValueCell(writer, intValue0);
                else WriteCellValue(writer, values[offset], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[1] is string style1) writer.Write(style1);
                if (values[offset + 1] is string stringValue1) WriteStringCellValue(writer, stringValue1, sharedStrings);
                else WriteCellValue(writer, values[offset + 1], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[2] is string style2) writer.Write(style2);
                if (values[offset + 2] is string stringValue2) WriteStringCellValue(writer, stringValue2, sharedStrings);
                else WriteCellValue(writer, values[offset + 2], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[3] is string style3) writer.Write(style3);
                if (values[offset + 3] is DateTime dateTimeValue3) WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateTimeValue3, dateSystem));
                else WriteCellValue(writer, values[offset + 3], DirectCellValueKind.DateTime, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[4] is string style4) writer.Write(style4);
                if (values[offset + 4] is double doubleValue4) WriteRawValueCell(writer, doubleValue4);
                else WriteCellValue(writer, values[offset + 4], DirectCellValueKind.Double, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[5] is string style5) writer.Write(style5);
                if (values[offset + 5] is int intValue5) WriteRawValueCell(writer, intValue5);
                else WriteCellValue(writer, values[offset + 5], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[6] is string style6) writer.Write(style6);
                if (values[offset + 6] is bool booleanValue6) writer.Write(booleanValue6 ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                else WriteCellValue(writer, values[offset + 6], DirectCellValueKind.Boolean, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("<c");
                if (styleAttributes?[7] is string style7) writer.Write(style7);
                if (values[offset + 7] is string stringValue7) WriteStringCellValue(writer, stringValue7, sharedStrings);
                else WriteCellValue(writer, values[offset + 7], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                writer.Write("</row>");
            }
            private static void WriteIntStringStringDateTimeDoubleIntBooleanStringCellValueRows(
                TextWriter writer,
                DirectCellValueRows cellValueRows,
                int rowCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                object?[] values = cellValueRows.Values;
                string prefix0 = cellReferencePrefixes[0];
                string prefix1 = cellReferencePrefixes[1];
                string prefix2 = cellReferencePrefixes[2];
                string prefix3 = cellReferencePrefixes[3];
                string prefix4 = cellReferencePrefixes[4];
                string prefix5 = cellReferencePrefixes[5];
                string prefix6 = cellReferencePrefixes[6];
                string prefix7 = cellReferencePrefixes[7];
                string? style0 = styleAttributes?[0];
                string? style1 = styleAttributes?[1];
                string? style2 = styleAttributes?[2];
                string? style3 = styleAttributes?[3];
                string? style4 = styleAttributes?[4];
                string? style5 = styleAttributes?[5];
                string? style6 = styleAttributes?[6];
                string? style7 = styleAttributes?[7];

                for (int sourceRowIndex = 0, offset = 0; sourceRowIndex < rowCount; sourceRowIndex++, offset += 8) {
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
                    if (style0 != null) writer.Write(style0);
                    if (values[offset] is int intValue0) WriteRawValueCell(writer, intValue0);
                    else WriteCellValue(writer, values[offset], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix1);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style1 != null) writer.Write(style1);
                    if (values[offset + 1] is string stringValue1) WriteStringCellValue(writer, stringValue1, sharedStrings);
                    else WriteCellValue(writer, values[offset + 1], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix2);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style2 != null) writer.Write(style2);
                    if (values[offset + 2] is string stringValue2) WriteStringCellValue(writer, stringValue2, sharedStrings);
                    else WriteCellValue(writer, values[offset + 2], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix3);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style3 != null) writer.Write(style3);
                    if (values[offset + 3] is DateTime dateTimeValue3) WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateTimeValue3, dateSystem));
                    else WriteCellValue(writer, values[offset + 3], DirectCellValueKind.DateTime, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix4);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style4 != null) writer.Write(style4);
                    if (values[offset + 4] is double doubleValue4) WriteRawValueCell(writer, doubleValue4);
                    else WriteCellValue(writer, values[offset + 4], DirectCellValueKind.Double, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix5);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style5 != null) writer.Write(style5);
                    if (values[offset + 5] is int intValue5) WriteRawValueCell(writer, intValue5);
                    else WriteCellValue(writer, values[offset + 5], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix6);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style6 != null) writer.Write(style6);
                    if (values[offset + 6] is bool booleanValue6) writer.Write(booleanValue6 ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                    else WriteCellValue(writer, values[offset + 6], DirectCellValueKind.Boolean, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write(prefix7);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style7 != null) writer.Write(style7);
                    if (values[offset + 7] is string stringValue7) WriteStringCellValue(writer, stringValue7, sharedStrings);
                    else WriteCellValue(writer, values[offset + 7], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);

                    writer.Write("</row>");
                    rowIndex++;
                }
            }
            private static bool IsIntStringStringDateTimeDoubleIntBooleanStringPlan(DirectCellValueKind[] cellValueKinds) {
                return cellValueKinds.Length == 8
                    && cellValueKinds[0] == DirectCellValueKind.Int32
                    && cellValueKinds[1] == DirectCellValueKind.String
                    && cellValueKinds[2] == DirectCellValueKind.String
                    && cellValueKinds[3] == DirectCellValueKind.DateTime
                    && cellValueKinds[4] == DirectCellValueKind.Double
                    && cellValueKinds[5] == DirectCellValueKind.Int32
                    && cellValueKinds[6] == DirectCellValueKind.Boolean
                    && cellValueKinds[7] == DirectCellValueKind.String;
            }

            private static void WriteIntStringStringDateTimeDoubleIntBooleanStringBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                string prefix0 = cellReferencePrefixes[0];
                string prefix1 = cellReferencePrefixes[1];
                string prefix2 = cellReferencePrefixes[2];
                string prefix3 = cellReferencePrefixes[3];
                string prefix4 = cellReferencePrefixes[4];
                string prefix5 = cellReferencePrefixes[5];
                string prefix6 = cellReferencePrefixes[6];
                string prefix7 = cellReferencePrefixes[7];
                string? style0 = styleAttributes?[0];
                string? style1 = styleAttributes?[1];
                string? style2 = styleAttributes?[2];
                string? style3 = styleAttributes?[3];
                string? style4 = styleAttributes?[4];
                string? style5 = styleAttributes?[5];
                string? style6 = styleAttributes?[6];
                string? style7 = styleAttributes?[7];

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] values = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");

                    writer.Write(prefix0);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style0 != null) writer.Write(style0);
                    if (values[0] is int intValue0) {
                        WriteRawValueCell(writer, intValue0);
                    } else {
                        WriteCellValue(writer, values[0], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix1);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style1 != null) writer.Write(style1);
                    if (values[1] is string stringValue1) {
                        WriteStringCellValue(writer, stringValue1, sharedStrings);
                    } else {
                        WriteCellValue(writer, values[1], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix2);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style2 != null) writer.Write(style2);
                    if (values[2] is string stringValue2) {
                        WriteStringCellValue(writer, stringValue2, sharedStrings);
                    } else {
                        WriteCellValue(writer, values[2], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix3);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style3 != null) writer.Write(style3);
                    if (values[3] is DateTime dateTimeValue3) {
                        WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateTimeValue3, dateSystem));
                    } else {
                        WriteCellValue(writer, values[3], DirectCellValueKind.DateTime, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix4);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style4 != null) writer.Write(style4);
                    if (values[4] is double doubleValue4) {
                        WriteRawValueCell(writer, doubleValue4);
                    } else {
                        WriteCellValue(writer, values[4], DirectCellValueKind.Double, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix5);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style5 != null) writer.Write(style5);
                    if (values[5] is int intValue5) {
                        WriteRawValueCell(writer, intValue5);
                    } else {
                        WriteCellValue(writer, values[5], DirectCellValueKind.Int32, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix6);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style6 != null) writer.Write(style6);
                    if (values[6] is bool booleanValue6) {
                        writer.Write(booleanValue6 ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                    } else {
                        WriteCellValue(writer, values[6], DirectCellValueKind.Boolean, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write(prefix7);
                    writer.Write(rowReference);
                    writer.Write('"');
                    if (style7 != null) writer.Write(style7);
                    if (values[7] is string stringValue7) {
                        WriteStringCellValue(writer, stringValue7, sharedStrings);
                    } else {
                        WriteCellValue(writer, values[7], DirectCellValueKind.String, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }
        }
    }
}
