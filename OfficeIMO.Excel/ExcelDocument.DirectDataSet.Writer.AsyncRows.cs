using System.IO.Compression;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static async Task<int> WriteAsyncWorksheetRows<T>(
                ZipArchive archive,
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
                var entry = archive.CreateEntry(
                    "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml",
                    CompressionLevel.Fastest);
                Stream? stream = null;
                PooledUtf8TextWriter? writer = null;
                bool completed = false;
                try {
                    stream = entry.Open();
                    writer = new PooledUtf8TextWriter(stream, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true);

                    writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                    if (sheet.StreamingLayout != null) writer.Write(sheet.StreamingLayout.ViewsXml);
                    WriteDeclaredColumns(writer, sheet.ColumnWidths, declaredStyles, sheet.StreamingLayout?.FirstColumn ?? 1);
                    writer.Write("<sheetData>");

                    int count = await WriteAsyncSheetData(writer, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        columnWritePlan, rows, writeRow, maximumRows, sharedStrings, declaredStyles, ct).ConfigureAwait(false);
                    writer.Write("</sheetData>");
                    WriteAsyncWorksheetEnd(writer, sheet);
                    await writer.FlushAsync().ConfigureAwait(false);
                    ct.ThrowIfCancellationRequested();
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
