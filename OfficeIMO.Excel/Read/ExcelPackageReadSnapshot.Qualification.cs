using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelPackageReadSnapshot {
        private static readonly XmlReaderSettings QualificationXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false
        };

        private readonly Dictionary<Uri, WorksheetExtent?> _worksheetExtents = new();

        /// <summary>
        /// Scans a worksheet on its first tabular read; a canceled scan is not kept. Only the chunk or
        /// schema sample the SDK route would actually buffer can stay there, where it reports the buffered
        /// cell limit. Dense reads check their own limit on either route, so sheet size alone does not
        /// disqualify a worksheet.
        /// </summary>
        internal bool QualifiesForTabularReaders(WorksheetPart worksheet, ExcelReaderOptions options) {
            WorksheetExtent? extent;
            lock (_gate) {
                if (!_worksheetExtents.TryGetValue(worksheet.Uri, out extent)) {
                    extent = ScanWorksheet(worksheet, options.CancellationToken);
                    _worksheetExtents[worksheet.Uri] = extent;
                }
            }

            int bufferedRows = Math.Max(
                Math.Min(ExcelSheetReader.DataReaderChunkRows, options.MaxDataReaderChunkRows),
                options.SchemaSampleRows);
            return extent is WorksheetExtent known
                && known.Columns * Math.Min(known.Rows, bufferedRows) <= options.MaxDataReaderBufferedCells;
        }

        private static WorksheetExtent? ScanWorksheet(WorksheetPart worksheet, CancellationToken cancellationToken) {
            try {
                using Stream stream = worksheet.GetStream(FileMode.Open, FileAccess.Read);
                using XmlReader reader = XmlReader.Create(stream, QualificationXmlReaderSettings);
                return ExcelSheetReader.MeasureCanonicalWorksheet(reader, cancellationToken);
            } catch (XmlException) {
                return null;
            } catch (IOException) {
                return null;
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
