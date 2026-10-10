using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>Enumerates the direct cells used by SDK worksheet projection and dimension discovery.</summary>
        private static IEnumerable<Cell> EnumerateOwnedSdkWorksheetCells(Worksheet worksheet) {
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) yield break;

            foreach (Row row in sheetData.Elements<Row>()) {
                foreach (Cell cell in row.Elements<Cell>()) {
                    yield return cell;
                }
            }
        }
    }
}
