using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            internal static async Task<int> WriteRowsAsync<T>(
                Stream stream,
                DirectDataSetWorkbookModel model,
                IAsyncEnumerable<T> rows,
                Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows,
                CancellationToken ct) {
                if (model.Sheets.Count != 1 || model.Sheets[0].HasTable) {
                    throw new InvalidOperationException(
                        "Asynchronous row exports require one package-native worksheet without an Excel table.");
                }

                DirectStylePlan stylePlan = DirectStylePlan.Create(model);
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, stylePlan, ct);
                using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
                WritePackagePreamble(archive, model, stylePlan, includeSharedStrings: false);

                return await WriteAsyncWorksheetRows(
                    archive,
                    model.Sheets[0],
                    model.DateTimeOffsetWriteStrategy,
                    model.DateSystem,
                    columnWritePlans[0],
                    rows,
                    writeRow,
                    maximumRows,
                    ct).ConfigureAwait(false);
            }

            private static async Task<int> WriteAsyncWorksheetRows<T>(
                ZipArchive archive,
                DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan,
                IAsyncEnumerable<T> rows,
                Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows,
                CancellationToken ct) {
                var entry = archive.CreateEntry(
                    "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml",
                    CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                WriteColumns(writer, sheet.ColumnWidths);
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
                            null,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings: null);
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
                    sharedStrings: null);

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
                } finally {
                    await enumerator.DisposeAsync().ConfigureAwait(false);
                }

                writer.Write("</sheetData></worksheet>");
                await writer.FlushAsync().ConfigureAwait(false);
                ct.ThrowIfCancellationRequested();
                return rowCount;
            }
        }
    }
}
