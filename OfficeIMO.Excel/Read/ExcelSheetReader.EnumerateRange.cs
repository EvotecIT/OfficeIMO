using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Range enumeration for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader
    {
        /// <summary>
        /// Enumerates non-empty cells within the given A1 range as typed values.
        /// </summary>
        public IEnumerable<CellValueInfo> EnumerateRange(string a1Range)
        {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) yield break;

            foreach (var row in sheetData.Elements<Row>())
            {
                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) break;

                foreach (var cell in row.Elements<Cell>())
                {
                    var (cIndex, _) = A1.ParseCellRef(cell.CellReference?.Value ?? string.Empty);
                    if (cIndex < c1 || cIndex > c2) continue;
                    var value = ConvertCell(cell);
                    if (value is not null || CellHasExplicitBlank(cell))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }
    }
}

