using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public class ExcelSheet {
        private readonly Sheet _sheet;
        public string Name;
        public UInt32Value Id;
        private readonly WorksheetPart _worksheetPart;

        public ExcelSheet(WorksheetPart worksheetPart) {

            _worksheetPart = worksheetPart;

        }

        public ExcelSheet(WorkbookPart workbookpart, SpreadsheetDocument spreadSheetDocument, string name) {
            UInt32Value id = 1;
            if (name == "") {
                name = "Sheet1";
            }
            
            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = id,
                Name = name
            };
            sheets.Append(sheet);

            this._sheet = sheet;
            this.Name = name;
            this.Id = sheet.SheetId;
            this._worksheetPart = worksheetPart;

        }
    }
}
