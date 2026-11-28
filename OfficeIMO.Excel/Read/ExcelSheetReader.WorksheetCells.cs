using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static IEnumerable<ReaderCell> EnumerateWorksheetCells(OfficeXmlNode worksheet) {
            foreach (OfficeXmlNode row in worksheet.Child(SpreadsheetXml, "sheetData").Children(SpreadsheetXml, "row")) {
                foreach (OfficeXmlNode cell in row.Children(SpreadsheetXml, "c")) {
                    yield return new ReaderCell(cell);
                }
            }
        }
    }
}
