using System.IO.Compression;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static async Task<int> WriteAsyncWorksheetRows<T>(
                ExcelPackageArchive archive,
                DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan,
                IAsyncEnumerable<T> rows,
                Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows,
                DirectSharedStringTable? sharedStrings,
                ExcelTabularStylePlan? declaredStyles,
                CancellationToken ct,
                Stylesheet? autoFitStyles = null) {
                ct.ThrowIfCancellationRequested();
                if (autoFitStyles != null) {
                    return await WriteAutoFitWorksheetRows(archive, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        columnWritePlan, rows, writeRow, maximumRows, sharedStrings, declaredStyles, autoFitStyles, ct).ConfigureAwait(false);
                }
                string entryName = "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml";
                Stream? stream = null;
                PooledUtf8TextWriter? writer = null;
                bool completed = false;
                try {
                    stream = archive.OpenEntry(entryName, CompressionLevel.Fastest, background: true);
                    // The row count is known only at the end. A rewritable destination gets a dimension placeholder that is
                    // filled in then, so readers can return rows before reaching the end of the worksheet.
                    byte[] head = StreamingWorksheetHead(dimension: null);
                    bool patchable = archive.TryWritePatchablePrefix(stream, head, out object? dimensionPrefix);
                    writer = new PooledUtf8TextWriter(stream, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true);
                    if (!patchable) writer.Write(StreamingWorksheetRoot);
                    if (sheet.StreamingLayout != null) writer.Write(sheet.StreamingLayout.ViewsXml);
                    WriteDeclaredColumns(writer, sheet.ColumnWidths, declaredStyles, sheet.StreamingLayout?.FirstColumn ?? 1);
                    writer.Write("<sheetData>");

                    int count = await WriteAsyncSheetData(writer, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        columnWritePlan, rows, writeRow, maximumRows, sharedStrings, declaredStyles, ct).ConfigureAwait(false);
                    writer.Write("</sheetData>");
                    WriteAsyncWorksheetEnd(writer, sheet);
                    writer.Flush();
                    writer.Dispose();
                    writer = null;
                    stream.Dispose();
                    stream = null;
                    ct.ThrowIfCancellationRequested();
                    if (patchable && GetStreamingDimension(sheet, count) is string dimension) {
                        archive.PatchPrefix(dimensionPrefix!, StreamingWorksheetHead(dimension));
                    }
                    completed = true;
                    return count;
                } finally {
                    try { writer?.Dispose(); }
                    catch when (!completed) {
                        // Keep the row or validation error when disposal also fails.
                    } finally {
                        try { stream?.Dispose(); }
                        catch when (!completed) { }
                    }
                }
            }

            private const string StreamingWorksheetRoot = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";

            // Wide enough for any range, "XFC1048575:XFD1048576"; shorter ranges are followed by spaces inside the tag.
            private const int StreamingDimensionWidth = 21;

            /// <summary>The worksheet start with a fixed-width dimension, so the placeholder and the final range have one length.</summary>
            private static byte[] StreamingWorksheetHead(string? dimension) {
                string reference = dimension ?? "A1";
                return Utf8NoBom.GetBytes(StreamingWorksheetRoot + "<dimension ref=\"" + reference + "\""
                    + new string(' ', StreamingDimensionWidth - reference.Length + 1) + "/>");
            }

            /// <summary>The range of the title, header and data rows a streamed table wrote, or null when it wrote none.</summary>
            private static string? GetStreamingDimension(DirectDataSetSheetModel sheet, int dataRows) {
                StreamingWorksheetLayout? layout = sheet.StreamingLayout;
                int rows = (layout?.Title != null ? 1 : 0) + (sheet.IncludeHeaders ? 1 : 0) + dataRows;
                if (rows == 0) return null;
                int firstRow = layout?.FirstRow ?? 1;
                int firstColumn = layout?.FirstColumn ?? 1;
                return ExcelA1.CellReference(firstRow, firstColumn) + ":"
                    + ExcelA1.CellReference(firstRow + rows - 1, firstColumn + sheet.Table.ColumnCount - 1);
            }

            private static async Task<int> WriteAsyncSheetData<T>(TextWriter writer, DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan, IAsyncEnumerable<T> rows, Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows, DirectSharedStringTable? sharedStrings, ExcelTabularStylePlan? declaredStyles, CancellationToken ct) {
                int columnCount = sheet.Table.ColumnCount;
                var layout = sheet.StreamingLayout;
                string[] cellReferencePrefixes = layout == null || layout.FirstColumn == 1 ? CreateCellReferencePrefixes(columnCount)
                    : Enumerable.Range(layout.FirstColumn, columnCount).Select(column => "<c r=\"" + ExcelA1.ColumnIndexToLetters(column)).ToArray();
                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                int rowIndex = layout?.FirstRow ?? 1;
                if (layout?.Title != null) {
                    string titleRow = InvariantNumberText.Get(rowIndex++);
                    writer.Write("<row r=\""); writer.Write(titleRow); writer.Write("\">");
                    WriteCell(writer, titleRow, cellReferencePrefixes[0], layout.Title, layout.TitleAttribute,
                        dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                    writer.Write("</row>");
                }
                if (sheet.IncludeHeaders) {
                    string headerRowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\""); writer.Write(headerRowReference); writer.Write("\">");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        WriteCell(
                            writer,
                            headerRowReference,
                            cellReferencePrefixes[columnIndex],
                            sheet.Table.GetColumnName(columnIndex),
                            layout?.HeaderAttributes[columnIndex] ?? declaredStyles?.Columns[columnIndex]?.Attribute,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings);
                    }
                    writer.Write("</row>");
                    rowIndex++;
                }

                var rowWriter = ExcelTabularRowWriter.Create(
                    writer,
                    rowIndex,
                    sheet.IncludeCellReferences,
                    cellReferencePrefixes,
                    styleAttributes,
                    valueStyleColumns,
                    sheet.UseCellValueNumberFormats,
                    dateTimeOffsetWriteStrategy,
                    dateSystem,
                    sharedStrings,
                    declaredStyles);

                return await WriteStreamingRowCallbacks(rows, rowWriter, writeRow, maximumRows, ct).ConfigureAwait(false);
            }

            private static void WriteAsyncWorksheetEnd(TextWriter writer, DirectDataSetSheetModel sheet) {
                if (sheet.HasTable) writer.Write("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>");
                writer.Write("</worksheet>");
            }
        }
    }
}
