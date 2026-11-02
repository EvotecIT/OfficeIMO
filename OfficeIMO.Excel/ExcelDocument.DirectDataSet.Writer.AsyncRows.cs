using System.IO.Compression;
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
                CancellationToken ct) {
                var entry = archive.CreateEntry(
                    "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml",
                    CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new PooledUtf8TextWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                WriteDeclaredColumns(writer, sheet.ColumnWidths, declaredStyles);
                writer.Write("<sheetData>");

                int columnCount = sheet.Table.ColumnCount;
                string[] cellReferencePrefixes = CreateCellReferencePrefixes(columnCount);
                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                int rowIndex = 1;
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        WriteCell(
                            writer,
                            headerRowReference,
                            cellReferencePrefixes[columnIndex],
                            sheet.Table.GetColumnName(columnIndex),
                            declaredStyles?.Columns[columnIndex]?.Attribute,
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

                int rowCount = 0;
                IAsyncEnumerator<T> enumerator = rows.GetAsyncEnumerator(ct);
                try {
                    while (true) {
                        ct.ThrowIfCancellationRequested();
                        if (!await enumerator.MoveNextAsync().ConfigureAwait(false)) {
                            break;
                        }
                        ct.ThrowIfCancellationRequested();
                        if (rowCount >= maximumRows) {
                            throw new InvalidOperationException(
                                "Asynchronous row export exceeds the maximum worksheet row count.");
                        }

                        rowWriter.BeginRow();
                        writeRow(rowWriter, enumerator.Current);
                        rowWriter.EndRow();
                        rowCount++;
                    }
                }
                finally {
                    await enumerator.DisposeAsync().ConfigureAwait(false);
                }

                writer.Write("</sheetData>");
                if (sheet.HasTable) writer.Write("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>");
                writer.Write("</worksheet>");
                await writer.FlushAsync().ConfigureAwait(false);
                ct.ThrowIfCancellationRequested();
                return rowCount;
            }
        }
    }
}
