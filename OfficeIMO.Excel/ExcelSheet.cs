﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single worksheet within an <see cref="ExcelDocument"/>.
    /// </summary>
    public class ExcelSheet {
        private readonly Sheet _sheet;

        /// <summary>
        /// Gets or sets the worksheet name.
        /// </summary>
        public string Name {
            get {
                return _sheet.Name;
            }
            set {
                _sheet.Name = value;
            }
        }
        private readonly UInt32Value Id;
        private readonly WorksheetPart _worksheetPart;
        private readonly SpreadsheetDocument _spreadSheetDocument;
        private readonly ExcelDocument _excelDocument;

        /// <summary>
        /// Initializes a worksheet from an existing <see cref="Sheet"/> element.
        /// </summary>
        /// <param name="excelDocument">Parent document.</param>
        /// <param name="spreadSheetDocument">Open XML spreadsheet document.</param>
        /// <param name="sheet">Underlying sheet element.</param>
        public ExcelSheet(ExcelDocument excelDocument, SpreadsheetDocument spreadSheetDocument, Sheet sheet) {
            _excelDocument = excelDocument;
            _sheet = sheet;
            _spreadSheetDocument = spreadSheetDocument;

            var list = _spreadSheetDocument.WorkbookPart.WorksheetParts.ToList();
            foreach (var worksheetPart in list) {
                var id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart);
                if (id == _sheet.Id) {
                    _worksheetPart = worksheetPart;
                }
            }
        }

        /// <summary>
        /// Creates a new worksheet and appends it to the workbook.
        /// </summary>
        /// <param name="excelDocument">Parent document.</param>
        /// <param name="workbookpart">Workbook part to add the worksheet to.</param>
        /// <param name="spreadSheetDocument">Open XML spreadsheet document.</param>
        /// <param name="name">Worksheet name.</param>
        public ExcelSheet(ExcelDocument excelDocument, WorkbookPart workbookpart, SpreadsheetDocument spreadSheetDocument, string name) {
            _excelDocument = excelDocument;
            _spreadSheetDocument = spreadSheetDocument;

            UInt32Value id = excelDocument.id.Max() + 1;
            if (name == "") {
                name = "Sheet1";
            }
            
            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = null;
            if (spreadSheetDocument.WorkbookPart.Workbook.Sheets != null) {
                sheets = spreadSheetDocument.WorkbookPart.Workbook.Sheets;
            } else {
                sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }

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

            excelDocument.id.Add(id);
        }
    }
}
