using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelPackageReadSnapshot {
        private const long NotCanonical = -1;
        private const long NotScanned = -2;
        private static readonly XmlReaderSettings QualificationXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false
        };

        private long _largestWorksheetCells = NotScanned;

        /// <summary>The first tabular read scans each worksheet once; a canceled scan is not kept.</summary>
        internal bool QualifiesForTabularReaders(ExcelReaderOptions options) {
            long largest;
            lock (_gate) {
                if (_largestWorksheetCells == NotScanned) {
                    _largestWorksheetCells = ScanWorksheets(options.CancellationToken);
                }
                largest = _largestWorksheetCells;
            }

            return largest != NotCanonical
                && largest <= Math.Min(options.MaxRangeCells, options.MaxDataReaderBufferedCells);
        }

        private long ScanWorksheets(CancellationToken cancellationToken) {
            WorkbookPart? workbookPart = Document.WorkbookPart;
            if (workbookPart == null) return NotCanonical;
            long largest = 0;
            foreach (WorksheetPart part in workbookPart.WorksheetParts) {
                try {
                    using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                    using XmlReader reader = XmlReader.Create(stream, QualificationXmlReaderSettings);
                    long cells = ExcelSheetReader.MeasureCanonicalWorksheet(reader, cancellationToken);
                    if (cells < 0) return NotCanonical;
                    largest = Math.Max(largest, cells);
                } catch (XmlException) {
                    return NotCanonical;
                } catch (InvalidDataException) {
                    return NotCanonical;
                }
            }

            return largest;
        }
    }
}
