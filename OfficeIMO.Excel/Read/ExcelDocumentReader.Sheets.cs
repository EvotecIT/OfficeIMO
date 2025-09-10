using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Additional sheet helpers for <see cref="ExcelDocumentReader"/>.
    /// </summary>
    public sealed partial class ExcelDocumentReader
    {
        /// <summary>
        /// Gets a reader by sheet index (1-based, Excel display order).
        /// </summary>
        public ExcelSheetReader GetSheet(int index)
        {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
            var list = GetSheetNames();
            if (index > list.Count) throw new ArgumentOutOfRangeException(nameof(index));

            var wb = _doc.WorkbookPart!.Workbook;
            var sheet = wb.Sheets!.Elements<Sheet>().ElementAt(index - 1);
            var wsPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheet.Id!);
            return new ExcelSheetReader(sheet.Name!, wsPart, _sst, _styles, _opt);
        }

        /// <summary>
        /// The number of worksheets in the workbook.
        /// </summary>
        public int SheetCount => _doc.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Count();
    }
}

