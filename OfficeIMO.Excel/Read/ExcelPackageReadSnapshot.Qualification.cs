using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelPackageReadSnapshot {
        private static readonly XmlReaderSettings QualificationXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false
        };

        private readonly Dictionary<string, WorksheetExtent?> _worksheetExtents = new(StringComparer.Ordinal);

        /// <summary>Checks whether the streamed projection preserves the worksheet's shape and the requested limits.</summary>
        internal bool QualifiesForTabularReaders(string worksheet, ExcelReaderOptions options) {
            WorksheetExtent? extent;
            lock (_gate) {
                if (!_worksheetExtents.TryGetValue(worksheet, out extent)) {
                    extent = MeasureWorksheet(worksheet, options);
                    _worksheetExtents.Add(worksheet, extent);
                }
            }
            return FitsBufferedWindow(extent, options);
        }

        private static bool FitsBufferedWindow(WorksheetExtent? extent, ExcelReaderOptions options) {
            int bufferedRows = Math.Max(Math.Min(ExcelSheetReader.DataReaderChunkRows, options.MaxDataReaderChunkRows),
                options.SchemaSampleRows);
            return extent is WorksheetExtent known
                && known.Columns * Math.Min(known.Rows, bufferedRows) <= options.MaxDataReaderBufferedCells;
        }

        private WorksheetExtent? MeasureWorksheet(string worksheet, ExcelReaderOptions options) {
            byte[]? buffer = null;
            try {
                const int maximumBufferedBytes = 256 * 1024 * 1024;
                if (PartBufferReader.TryGetLength(worksheet, out long length) && length >= 1_048_576
                    && PartBufferReader.TryRead(worksheet, maximumBufferedBytes, options.CancellationToken,
                        out buffer, out int bytesRead, ExcelSheetReader.CanBufferUtf8WorksheetPrefix)) {
                    WorksheetExtent? extent;
                    bool validatedIndexedXml = false;
                    if (ExcelSheetReader.TryMeasureCanonicalWorksheet(buffer!, bytesRead, options.Execution, options.CancellationToken,
                        out WorksheetExtent measured)) {
                        extent = measured;
                    } else {
                        using var buffered = new MemoryStream(buffer!, 0, bytesRead, writable: false);
                        using XmlReader bufferedReader = XmlReader.Create(buffered, QualificationXmlReaderSettings);
                        extent = ExcelSheetReader.MeasureCanonicalWorksheet(bufferedReader, options.CancellationToken, out validatedIndexedXml);
                    }
                    if (FitsBufferedWindow(extent, options)
                        && PartBufferReader.TryRetainReadBuffer(worksheet, buffer!, bytesRead, validatedIndexedXml)) {
                        buffer = null;
                    }
                    return extent;
                }

                using Stream stream = PartBufferReader.OpenPart(worksheet, int.MaxValue, options.CancellationToken);
                using XmlReader reader = XmlReader.Create(stream, QualificationXmlReaderSettings);
                return ExcelSheetReader.MeasureCanonicalWorksheet(reader, options.CancellationToken, out _);
            } catch (XmlException) {
                return null;
            } catch (IOException exception) when (!ExcelReadLimitFailure.Is(exception)) {
                return null;
            } finally {
                if (buffer != null) OpenXmlPartBufferPool.Return(buffer);
            }
        }
    }

    internal readonly struct WorksheetExtent {
        internal WorksheetExtent(long rows, long columns) {
            Rows = rows;
            Columns = columns;
        }

        internal long Rows { get; }

        internal long Columns { get; }
    }
}
