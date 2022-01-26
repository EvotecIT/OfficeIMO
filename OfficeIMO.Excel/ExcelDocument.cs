using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable {
        public SpreadsheetDocument _spreadsheetDocument;
        private WorkbookPart _workBookPart;
        public string FilePath;

        public static ExcelDocument Create(string filePath) {
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            
            document._spreadsheetDocument = spreadSheetDocument;

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            document._workBookPart = workbookpart;

            return document;
        }

        public static ExcelDocument Create(string filePath, string workSheetName) {
            ExcelDocument excelDocument = Create(filePath);
            excelDocument.AddWorkSheet(workSheetName);
            return excelDocument; 
        }

        public ExcelSheet AddWorkSheet(string workSheetName = "") {
            ExcelSheet excelSheet = new ExcelSheet(_workBookPart, _spreadsheetDocument);

            return excelSheet;
        }
        
        public void Open(string filePath = "", bool openExcel = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }
            Helpers.Open(filePath, openExcel);
        }

        public void Close() {
            this._spreadsheetDocument.Close();
        }

        public void Save(string filePath, bool openExcel) {
            this._workBookPart.Workbook.Save();

            this.Open(filePath, openExcel);
        }

        public void Save() {
            this.Save("", false);
        }

        public void Save(bool openExcel) {
            this.Save("", openExcel);
        }

        public void Dispose() {
            if (this._spreadsheetDocument != null) {
                this._spreadsheetDocument.Dispose();
            }
        }
    }
}
