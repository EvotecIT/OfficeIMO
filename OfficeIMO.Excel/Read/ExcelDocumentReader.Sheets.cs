using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Additional sheet helpers for <see cref="ExcelDocumentReader"/>.
    /// </summary>
    internal sealed partial class ExcelDocumentReader {
        /// <summary>
        /// Gets a reader by sheet index (1-based, Excel display order).
        /// </summary>
        public ExcelSheetReader GetSheet(int index) {
            ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
            if (TryGetSheetByIndexXmlFast(index, out string fastSheetName, out WorksheetPart fastWorksheetPart)) {
                return CreateSheetReader(fastSheetName, fastWorksheetPart);
            }

            var wb = WorkbookRoot;
            Sheet? sheet = null;
            int currentIndex = 0;
            foreach (var candidate in wb.Sheets!.Elements<Sheet>()) {
                if (!TryGetWorksheetPart(candidate, out _)) {
                    continue;
                }

                currentIndex++;
                if (currentIndex == index) {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet is null) throw new ArgumentOutOfRangeException(nameof(index));
            if (!TryGetWorksheetPart(sheet, out WorksheetPart? wsPart)) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return CreateSheetReader(sheet.Name!, wsPart!);
        }

        /// <summary>
        /// The number of worksheets in the workbook.
        /// </summary>
        public int SheetCount {
            get {
                if (TryGetSheetCountXmlFast(out int fastCount)) {
                    return fastCount;
                }

                int count = 0;
                foreach (var sheet in WorkbookRoot.Sheets!.Elements<Sheet>()) {
                    if (TryGetWorksheetPart(sheet, out _)) {
                        count++;
                    }
                }

                return count;
            }
        }
    }
}

