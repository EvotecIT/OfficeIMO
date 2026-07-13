using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private class CellUpdate {
            public int Row { get; }
            public int Column { get; }
            public string Text { get; }
            public DocumentFormat.OpenXml.Spreadsheet.CellValues DataType { get; }
            public bool IsSharedString { get; }

            public CellUpdate(int row, int column, string text, DocumentFormat.OpenXml.Spreadsheet.CellValues dataType, bool isSharedString) {
                Row = row;
                Column = column;
                Text = text;
                DataType = dataType;
                IsSharedString = isSharedString;
            }
        }

        private CellUpdate PrepareCellUpdate(int row, int column, object value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, dataType) = CoerceValueHelper.Coerce(
                value,
                s => new CellValue(s),
                dateTimeOffsetStrategy,
                _excelDocument.DateSystem);

            bool isSharedString = dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            return new CellUpdate(row, column, cellValue.Text ?? string.Empty, dataType, isSharedString);
        }

        private void ApplyCellUpdate(CellUpdate update) {
            Cell cell = GetCell(update.Row, update.Column);
            if (update.IsSharedString) {
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(update.Text);
                cell.CellValue = new CellValue(SharedStringIndexText.Get(sharedStringIndex));
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            } else {
                cell.CellValue = new CellValue(update.Text);
                cell.DataType = update.DataType;
            }
        }

        private void ValidateWorksheetXml()
            => WorksheetIntegrityValidator.Validate(_worksheetPart, EffectiveExecution, Name);
    }
}
