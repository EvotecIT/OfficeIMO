using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable {
        internal List<UInt32Value> id = new List<UInt32Value>() { 0 };

        public List<ExcelSheet> Sheets {
            get {
                List<ExcelSheet> listExcel = new List<ExcelSheet>();
                if (_spreadSheetDocument.WorkbookPart.Workbook.Sheets != null) {
                    var elements = _spreadSheetDocument.WorkbookPart.Workbook.Sheets.OfType<Sheet>().ToList();
                    foreach (Sheet s in elements) {
                        ExcelSheet excelSheet = new ExcelSheet(this, _spreadSheetDocument, s);
                        id.Add(s.SheetId);
                        listExcel.Add(excelSheet);
                    }
                }

                return listExcel;
            }
        }

        public SpreadsheetDocument _spreadSheetDocument;
        private WorkbookPart _workBookPart;
        public string FilePath;

        public static ExcelDocument Create(string filePath) {
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            document._spreadSheetDocument = spreadSheetDocument;

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            document._workBookPart = workbookpart;

            return document;
        }
        public static ExcelDocument Load(string filePath, bool readOnly = false, bool autoSave = false) {
            if (filePath != null) {
                if (!File.Exists(filePath)) {
                    throw new FileNotFoundException("File doesn't exists", filePath);
                }
            }
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            FileMode fileMode = readOnly ? FileMode.Open : FileMode.OpenOrCreate;
            var package = Package.Open(filePath, fileMode);

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(package, openSettings);

            document._spreadSheetDocument = spreadSheetDocument;

            //// Add a WorkbookPart to the document.
            document._workBookPart = document._spreadSheetDocument.WorkbookPart;

            return document;
        }

        public static ExcelDocument Create(string filePath, string workSheetName) {
            ExcelDocument excelDocument = Create(filePath);
            excelDocument.AddWorkSheet(workSheetName);
            return excelDocument;
        }

        public ExcelSheet AddWorkSheet(string workSheetName = "") {
            ExcelSheet excelSheet = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, workSheetName);

            return excelSheet;
        }

        public void Open(string filePath = "", bool openExcel = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }
            Helpers.Open(filePath, openExcel);
        }

        public void Close() {
            this._spreadSheetDocument.Dispose();
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
            if (this._spreadSheetDocument != null) {
                this._spreadSheetDocument.Dispose();
            }
        }
    }
}
