using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("style")]
        [InlineData("boolean")]
        [InlineData("order")]
        [InlineData("row")]
        [InlineData("none")]
        public void Column_edit_diagnostics_match_a_full_validation(string defect) {
            using var stream = new MemoryStream();
            using (var created = ExcelDocument.Create(stream)) {
                ExcelSheet source = created.AddSheet("Data");
                for (int row = 1; row <= 40; row++) {
                    source.Cell(row, 1).SetValue(row);
                    source.Cell(row, 2).SetFormula("A" + row + "*2");
                    source.Cell(row, 3).SetValue("Text");
                }
                created.Save(stream);
            }

            stream.Position = 0;
            using var document = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            ExcelSheet sheet = document.Sheets[0];
            SheetData data = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!;
            Row last = data.Elements<Row>().Last();
            Cell target = last.Elements<Cell>().First();
            switch (defect) {
                case "style": target.StyleIndex = 900U; break;
                case "boolean":
                    target.DataType = CellValues.Boolean;
                    target.CellValue = new CellValue("7");
                    break;
                case "order":
                    Cell formulaCell = last.Elements<Cell>().ElementAt(1);
                    formulaCell.RemoveAllChildren();
                    formulaCell.AppendChild(new CellValue("1"));
                    formulaCell.AppendChild(new CellFormula("1+1"));
                    break;
                case "row": last.RowIndex = 2_000_000U; break;
            }

            ExcelMutationResult result = sheet.InsertColumnsTransactional(2);

            string[] expected = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365)
                .Validate(document.OpenXmlDocument).Select(error => error.Description ?? string.Empty).ToArray();
            Assert.Equal(defect == "none", expected.Length == 0);
            Assert.Equal(expected, result.Diagnostics.Select(item => item.Message).ToArray());
            Assert.Same(data, document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>());
        }
    }
}
