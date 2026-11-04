using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // XLBench template scenario: rows written below a template's header must keep its column formats.
        [Fact]
        public void NewCellsInStyledRowsAndColumnsStartWithThatStyle() {
            using var template = new MemoryStream();
            using (var package = SpreadsheetDocument.Create(template, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbookPart = package.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets(new Sheet { Name = "Report", SheetId = 1U, Id = "rId1" }));
                WorkbookStylesPart styles = workbookPart.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(
                    new NumberingFormats(new NumberingFormat { NumberFormatId = 164U, FormatCode = "yyyy\\-mm\\-dd" }) { Count = 1U },
                    new Fonts(new Font()) { Count = 1U },
                    new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1U },
                    new Borders(new Border()) { Count = 1U },
                    new CellStyleFormats(new CellFormat()) { Count = 1U },
                    new CellFormats(
                        new CellFormat(),
                        new CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true },
                        new CellFormat { NumberFormatId = 4U, ApplyNumberFormat = true },
                        new CellFormat { NumberFormatId = 10U, ApplyNumberFormat = true }) { Count = 4U });
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
                worksheetPart.Worksheet = new Worksheet(
                    new Columns(
                        new Column { Min = 1U, Max = 1U, Width = 12D, Style = 1U, CustomWidth = true },
                        new Column { Min = 2U, Max = 3U, Width = 12D, Style = 2U, CustomWidth = true }),
                    new SheetData(
                        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Date")) }) { RowIndex = 1U },
                        new Row { RowIndex = 3U, StyleIndex = 3U, CustomFormat = true }));
            }

            template.Position = 0;
            using var document = ExcelDocument.Load(template);
            ExcelSheet sheet = document["Report"];
            sheet.Cell("A5").SetValue(new DateTime(2026, 1, 31));
            sheet.Cell("B5").SetValue(12.5m);
            sheet.Cell("C5").SetFormula("B5*2");
            sheet.Cell("D5").SetValue("unstyled column");
            sheet.Cell("B3").SetValue(0.25);

            Assert.Equal("yyyy\\-mm\\-dd", sheet.Cell("A5").GetStyle().NumberFormatCode);
            Assert.Equal(4U, sheet.Cell("B5").GetStyle().NumberFormatId);
            Assert.Equal(4U, sheet.Cell("C5").GetStyle().NumberFormatId);
            Assert.Equal(0U, sheet.Cell("D5").GetStyle().NumberFormatId);
            Assert.Equal(10U, sheet.Cell("B3").GetStyle().NumberFormatId);
        }
    }
}
