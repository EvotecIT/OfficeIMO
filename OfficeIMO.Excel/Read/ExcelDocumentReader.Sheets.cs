using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Additional sheet helpers for <see cref="ExcelDocumentReader"/>.
    /// </summary>
    public sealed partial class ExcelDocumentReader {
        /// <summary>
        /// Gets a reader by sheet index (1-based, Excel display order).
        /// </summary>
        public ExcelSheetReader GetSheet(int index) {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
            if (TryGetSheetByIndexXmlFast(index, out string fastSheetName, out WorksheetPart fastWorksheetPart)) {
                return new ExcelSheetReader(fastSheetName, fastWorksheetPart, _sst, _styles, _opt, _dateSystem, _owns);
            }

            var wb = WorkbookRoot;
            Sheet? sheet = null;
            int currentIndex = 0;
            foreach (var candidate in wb.Sheets!.Elements<Sheet>()) {
                currentIndex++;
                if (currentIndex == index) {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet is null) throw new ArgumentOutOfRangeException(nameof(index));
            var wsPart = (WorksheetPart)WorkbookPartRoot.GetPartById(sheet.Id!);
            return new ExcelSheetReader(sheet.Name!, wsPart, _sst, _styles, _opt, _dateSystem, _owns);
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
                foreach (var _ in WorkbookRoot.Sheets!.Elements<Sheet>()) {
                    count++;
                }

                return count;
            }
        }
    }
}

