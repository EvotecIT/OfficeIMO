using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            /// <summary>Spools only sheetData so final column widths precede it without retaining source rows.</summary>
            private static async Task<int> WriteAutoFitWorksheetRows<T>(ZipArchive archive, DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan, IAsyncEnumerable<T> rows, Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows, DirectSharedStringTable? sharedStrings, ExcelTabularStylePlan? declaredStyles,
                Stylesheet stylesheet, CancellationToken ct) {
                string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-" + Guid.NewGuid().ToString("N") + ".xml");
                using var spool = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                    64 * 1024, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
                int count;
                using (var bodyWriter = new PooledUtf8TextWriter(spool, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true)) {
                    bodyWriter.Write("<sheetData xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    count = await WriteAsyncSheetData(bodyWriter, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        columnWritePlan, rows, writeRow, maximumRows, sharedStrings, declaredStyles, ct).ConfigureAwait(false);
                    bodyWriter.Write("</sheetData>");
                    await bodyWriter.FlushAsync().ConfigureAwait(false);
                }

                spool.Position = 0;
                double[] widths = MeasureStreamingColumns(spool, sheet, stylesheet, dateSystem, ct);
                spool.Position = 0;
                var entry = archive.CreateEntry("xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using (var writer = new PooledUtf8TextWriter(stream, Utf8NoBom, XmlWriterBufferSize)) {
                    writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                    WriteDeclaredColumns(writer, widths, declaredStyles);
                    await writer.FlushAsync().ConfigureAwait(false);
                    await spool.CopyToAsync(stream, 64 * 1024, ct).ConfigureAwait(false);
                    WriteAsyncWorksheetEnd(writer, sheet);
                    await writer.FlushAsync().ConfigureAwait(false);
                }
                ct.ThrowIfCancellationRequested();
                return count;
            }

            private static double[] MeasureStreamingColumns(Stream spool, DirectDataSetSheetModel sheet,
                Stylesheet stylesheet, ExcelDateSystem dateSystem, CancellationToken ct) {
                var widths = new double[sheet.Table.ColumnCount];
                var measurer = new ExcelSheet.AutoFitWidthMeasurer(stylesheet);
                using var reader = XmlReader.Create(spool, new XmlReaderSettings { CloseInput = false, DtdProcessing = DtdProcessing.Prohibit });
                int row = 0;
                int column = 0;
                while (!reader.EOF) {
                    ct.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row") { row++; column = 0; }
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c") {
                        var cell = new Cell(reader.ReadOuterXml());
                        double width = measurer.MeasureCell(cell, dateSystem,
                            tableHeader: row == 1 && sheet.IncludeHeaders && sheet.HasTable,
                            filterButton: sheet.IncludeAutoFilter);
                        widths[column] = Math.Max(widths[column], NormalizeColumnWidth(width));
                        column++;
                    } else {
                        reader.Read();
                    }
                }
                return widths;
            }
        }
    }
}
