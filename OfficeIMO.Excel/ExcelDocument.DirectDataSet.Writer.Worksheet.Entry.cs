using System.IO.Compression;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            /// <summary>Joins the worksheet byte compressor before closing its ZIP entry.</summary>
            private static void WriteWorksheet(ExcelPackageArchive archive, DirectDataSetSheetModel sheet,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings, DirectStylePlan stylePlan,
                DirectColumnWritePlan columnWritePlan, CancellationToken ct, string? worksheetPath = null,
                string? tableRelationshipId = null, bool allowBackgroundCompression = false) {
                string entryName = worksheetPath ?? "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml";
                Stream? stream = null;
                PooledUtf8TextWriter? writer = null;
                StreamingWorksheetCompression? compression = null;
                bool completed = false;
                try {
                    stream = archive.OpenEntry(entryName, CompressionLevel.Fastest);
                    if (allowBackgroundCompression) compression = new StreamingWorksheetCompression(stream);
                    writer = new PooledUtf8TextWriter(compression ?? stream, Utf8NoBom, XmlWriterBufferSize, leaveOpen: true);
                    WriteWorksheetContents(writer, sheet, dateTimeOffsetWriteStrategy, dateSystem,
                        sharedStrings, stylePlan, columnWritePlan, ct, tableRelationshipId);
                    writer.Dispose();
                    writer = null;
                    compression?.Complete();
                    completed = true;
                } finally {
                    try { writer?.Dispose(); }
                    catch when (!completed) {
                        // Retain the producer's error if flushing also fails.
                    } finally {
                        try { compression?.Dispose(); }
                        finally {
                            try { stream?.Dispose(); }
                            catch when (!completed) {
                                // The entry cleanup must not replace the original failure.
                            }
                        }
                    }
                }
            }
        }
    }
}
