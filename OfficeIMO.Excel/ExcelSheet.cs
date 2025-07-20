using System;
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

        private Stylesheet Stylesheet {
            get {
                if (_spreadSheetDocument.WorkbookPart.WorkbookStylesPart == null) {
                    var styles = _spreadSheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    styles.Stylesheet = new Stylesheet(
                        new Fonts(new Font()) { Count = 1 },
                        new Fills(
                            new Fill(new PatternFill { PatternType = PatternValues.None }),
                            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 },
                        new Borders(new Border()) { Count = 1 },
                        new CellFormats(new CellFormat()) { Count = 1 }
                    );
                    styles.Stylesheet.Save();
                }
                return _spreadSheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            }
        }

        private uint AddStyle(Font? font, Fill? fill, Border? border, string? numberFormat) {
            var stylesheet = Stylesheet;

            var fonts = stylesheet.Fonts ??= new Fonts(new Font()) { Count = 1 };
            uint fontId;
            if (font != null) {
                fonts.Append((Font)font.CloneNode(true));
                fonts.Count = (uint)fonts.ChildElements.Count;
                fontId = (uint)(fonts.ChildElements.Count - 1);
            } else {
                fontId = 0;
            }

            var fills = stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 };
            uint fillId;
            if (fill != null) {
                fills.Append((Fill)fill.CloneNode(true));
                fills.Count = (uint)fills.ChildElements.Count;
                fillId = (uint)(fills.ChildElements.Count - 1);
            } else {
                fillId = 0;
            }

            var borders = stylesheet.Borders ??= new Borders(new Border()) { Count = 1 };
            uint borderId;
            if (border != null) {
                borders.Append((Border)border.CloneNode(true));
                borders.Count = (uint)borders.ChildElements.Count;
                borderId = (uint)(borders.ChildElements.Count - 1);
            } else {
                borderId = 0;
            }

            uint numFmtId = 0;
            if (!string.IsNullOrEmpty(numberFormat)) {
                var numFmts = stylesheet.NumberingFormats ??= new NumberingFormats();
                numFmtId = (uint)(164 + numFmts.ChildElements.Count);
                numFmts.Append(new NumberingFormat { NumberFormatId = numFmtId, FormatCode = numberFormat });
                numFmts.Count = (uint)numFmts.ChildElements.Count;
            }

            var cellFormats = stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };

            var cf = new CellFormat {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                NumberFormatId = numFmtId
            };
            if (numFmtId != 0) cf.ApplyNumberFormat = true;
            cf.ApplyBorder = true;
            cf.ApplyFont = true;
            cf.ApplyFill = true;
            cellFormats.Append(cf);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            stylesheet.Save();

            return (uint)cellFormats.ChildElements.Count - 1;
        }

        internal Font? GetFont(ExcelCell cell) {
            if (cell._cell.StyleIndex == null) return null;
            var stylesheet = Stylesheet;
            var format = (CellFormat)stylesheet.CellFormats.ChildElements[(int)cell._cell.StyleIndex.Value];
            return (Font)stylesheet.Fonts.ChildElements[(int)format.FontId.Value];
        }

        internal Fill? GetFill(ExcelCell cell) {
            if (cell._cell.StyleIndex == null) return null;
            var stylesheet = Stylesheet;
            var format = (CellFormat)stylesheet.CellFormats.ChildElements[(int)cell._cell.StyleIndex.Value];
            return (Fill)stylesheet.Fills.ChildElements[(int)format.FillId.Value];
        }

        internal Border? GetBorder(ExcelCell cell) {
            if (cell._cell.StyleIndex == null) return null;
            var stylesheet = Stylesheet;
            var format = (CellFormat)stylesheet.CellFormats.ChildElements[(int)cell._cell.StyleIndex.Value];
            return (Border)stylesheet.Borders.ChildElements[(int)format.BorderId.Value];
        }

        internal string? GetNumberFormat(ExcelCell cell) {
            if (cell._cell.StyleIndex == null) return null;
            var stylesheet = Stylesheet;
            var format = (CellFormat)stylesheet.CellFormats.ChildElements[(int)cell._cell.StyleIndex.Value];
            if (format.NumberFormatId != null && stylesheet.NumberingFormats != null) {
                var nFmt = stylesheet.NumberingFormats.Elements<NumberingFormat>().FirstOrDefault(n => n.NumberFormatId.Value == format.NumberFormatId.Value);
                return nFmt?.FormatCode?.Value;
            }
            return null;
        }

        internal void ApplyStyle(ExcelCell cell, Font? font, Fill? fill, Border? border, string? numberFormat) {
            var styleIndex = AddStyle(font, fill, border, numberFormat);
            cell._cell.StyleIndex = styleIndex;
        }

        public ExcelCell GetCell(string cellReference) {
            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            uint rowIndex = uint.Parse(new string(cellReference.Where(char.IsDigit).ToArray()));
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null) {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
            if (cell == null) {
                cell = new Cell { CellReference = cellReference };
                row.Append(cell);
            }
            return new ExcelCell(this, cell);
        }

        public void MergeCells(string reference) {
            var mergeCells = _worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null) {
                mergeCells = new MergeCells();
                var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
                sheetData.InsertAfterSelf(mergeCells);
            }
            mergeCells.Append(new MergeCell { Reference = reference });
        }

        public void UnmergeCells(string reference) {
            var mergeCells = _worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null) return;
            var mc = mergeCells.Elements<MergeCell>().FirstOrDefault(m => m.Reference?.Value == reference);
            mc?.Remove();
        }
    }
}
