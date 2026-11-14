using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            /// <summary>Sheet data up to this size is spooled in memory; a larger sheet moves to a temporary file.</summary>
            private const int AutoFitSpoolMemoryLimit = 16 * 1024 * 1024;

            /// <summary>Spools only sheetData so final column widths precede it without retaining source rows.</summary>
            private static async Task<int> WriteAutoFitWorksheetRows<T>(ExcelPackageArchive archive, DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan, IAsyncEnumerable<T> rows, Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows, DirectSharedStringTable? sharedStrings, ExcelTabularStylePlan? declaredStyles,
                Stylesheet stylesheet, CancellationToken ct) {
                using var spool = new ExcelSpillStream(AutoFitSpoolMemoryLimit, "OfficeIMO-Excel-AutoFit-");
                int count;
                using (var bodyWriter = new PooledUtf8TextWriter(spool, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true)) {
                    bodyWriter.Write("<sheetData xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    count = await WriteAsyncSheetData(bodyWriter, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        columnWritePlan, rows, writeRow, maximumRows, sharedStrings, declaredStyles, ct).ConfigureAwait(false);
                    bodyWriter.Write("</sheetData>");
                    bodyWriter.Flush();
                }

                spool.Position = 0;
                double[] widths = MeasureStreamingColumns(spool, sheet, stylesheet, dateSystem, ct);
                spool.Position = 0;
                using var stream = archive.OpenEntry("xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
                using (var writer = new PooledUtf8TextWriter(stream, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true)) {
                    writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                    if (sheet.StreamingLayout != null) writer.Write(sheet.StreamingLayout.ViewsXml);
                    WriteDeclaredColumns(writer, widths, declaredStyles, sheet.StreamingLayout?.FirstColumn ?? 1);
                    writer.Flush();
                    await spool.CopyToAsync(stream, 64 * 1024, ct).ConfigureAwait(false);
                    WriteAsyncWorksheetEnd(writer, sheet);
                    writer.Flush();
                }
                ct.ThrowIfCancellationRequested();
                return count;
            }

            /// <summary>
            /// Measures the spooled cells. Cells are read from the XML reader without building OpenXml cells; shared strings
            /// and unknown types are measured from their markup, and content the reader does not expect sends the whole
            /// sheet through the markup path, so every cell is measured the same way either way.
            /// </summary>
            private static double[] MeasureStreamingColumns(Stream spool, DirectDataSetSheetModel sheet,
                Stylesheet stylesheet, ExcelDateSystem dateSystem, CancellationToken ct, bool cellsFromMarkup = false) {
                var widths = new double[sheet.Table.ColumnCount];
                var measurer = new ExcelSheet.AutoFitWidthMeasurer(stylesheet);
                long start = spool.Position;
                if (!cellsFromMarkup && TryMeasureSpooledCells(spool, sheet, measurer, dateSystem, widths, ct)) return widths;
                spool.Position = start;
                Array.Clear(widths, 0, widths.Length);
                using (var reader = XmlReader.Create(spool, new XmlReaderSettings { CloseInput = false, DtdProcessing = DtdProcessing.Prohibit })) {
                    int row = 0;
                    int column = 0;
                    while (!reader.EOF) {
                        ct.ThrowIfCancellationRequested();
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row") {
                            row = int.TryParse(reader.GetAttribute("r"), out int rowNumber) ? rowNumber : row + 1;
                            column = 0;
                        }
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c") {
                            bool tableHeader = row == (sheet.StreamingLayout?.HeaderRow ?? 1) && sheet.IncludeHeaders && sheet.HasTable;
                            ExcelSheet.AutoFitCellParts parts = default;
                            var read = cellsFromMarkup ? ExcelSheet.AutoFitCellParts.ReadResult.Markup
                                : ExcelSheet.AutoFitCellParts.Read(reader, out parts);
                            if (read == ExcelSheet.AutoFitCellParts.ReadResult.Unsupported) {
                                cellsFromMarkup = true;
                                break;
                            }

                            double width = read == ExcelSheet.AutoFitCellParts.ReadResult.Read
                                ? measurer.MeasureCell(parts, dateSystem, tableHeader, sheet.IncludeAutoFilter)
                                : measurer.MeasureCell(new Cell(reader.ReadOuterXml()), dateSystem, tableHeader, sheet.IncludeAutoFilter);
                            if (sheet.StreamingLayout == null || sheet.StreamingLayout.AutoFitColumns[column])
                                widths[column] = Math.Max(widths[column], NormalizeColumnWidth(width));
                            column++;
                        } else {
                            reader.Read();
                        }
                    }

                    if (!cellsFromMarkup || reader.EOF) return widths;
                }

                spool.Position = start;
                return MeasureStreamingColumns(spool, sheet, stylesheet, dateSystem, ct, cellsFromMarkup: true);
            }

            /// <summary>
            /// Measures the spooled cells from their bytes, which avoids parsing the markup a second time. Returns false
            /// for markup the scanner does not read, and the caller then measures every cell through the XML reader.
            /// </summary>
            private static bool TryMeasureSpooledCells(Stream spool, DirectDataSetSheetModel sheet, ExcelSheet.AutoFitWidthMeasurer measurer,
                ExcelDateSystem dateSystem, double[] widths, CancellationToken ct) {
                var scanner = new ExcelSheet.AutoFitSpoolScanner(spool);
                int headerRow = sheet.StreamingLayout?.HeaderRow ?? 1;
                bool[]? measured = sheet.StreamingLayout?.AutoFitColumns;
                int row = 0;
                int column = 0;
                int cells = 0;
                while (true) {
                    switch (scanner.Next(out int rowNumber, out ExcelSheet.AutoFitCellParts parts)) {
                        case ExcelSheet.AutoFitSpoolScanner.Token.End:
                            return true;
                        case ExcelSheet.AutoFitSpoolScanner.Token.Unsupported:
                            return false;
                        case ExcelSheet.AutoFitSpoolScanner.Token.Row:
                            row = rowNumber > 0 ? rowNumber : row + 1;
                            column = 0;
                            if ((++cells & 0xFF) == 0) ct.ThrowIfCancellationRequested();
                            break;
                        default:
                            if (column >= widths.Length) return false;
                            bool tableHeader = row == headerRow && sheet.IncludeHeaders && sheet.HasTable;
                            double width = measurer.MeasureCell(parts, dateSystem, tableHeader, sheet.IncludeAutoFilter);
                            if (measured == null || measured[column]) widths[column] = Math.Max(widths[column], NormalizeColumnWidth(width));
                            column++;
                            break;
                    }
                }
            }
        }
    }
}
