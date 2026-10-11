using System.Data.Common;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("shared-string", "worksheet")]
        [InlineData("style", "worksheet")]
        [InlineData("formula", "worksheet")]
        [InlineData("shared-string", "row")]
        [InlineData("style", "row")]
        [InlineData("formula", "row")]
        [InlineData("shared-string", "cell")]
        [InlineData("style", "cell")]
        [InlineData("formula", "cell")]
        public void DataReader_BorrowedSdkIgnoresExtensionOwnedCellReferences(string kind, string placement) {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Number");
            sheet.CellValue(2, 1, 1);
            document.Save();
            Worksheet worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var payload = new OpenXmlUnknownElement("e", "metadata", "urn:officeimo:cell-ownership");
            payload.Append(CreateSdkOwnershipCell(kind));
            worksheet.AddNamespaceDeclaration("e", "urn:officeimo:cell-ownership");
            worksheet.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "e" };
            if (placement == "worksheet") {
                worksheet.Append(new WorksheetExtensionList(new WorksheetExtension(payload) {
                    Uri = "{564B230A-5555-4000-8000-000000000001}"
                }));
            } else {
                Row row = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Last();
                if (placement == "row") row.Append(payload);
                else row.Elements<Cell>().First().Append(payload);
            }

            using DbDataReader reader = document.CreateDataReader(new ExcelReadOptions { UseCachedFormulaResult = false });
            Assert.Equal(1, reader.FieldCount);
            Assert.Equal("Number", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.False(reader.Read());
            reader.Close();
            sheet.CellValue(3, 1, 2);
            Assert.NotNull(document.OpenXmlDocument.WorkbookPart);
        }

        [Fact]
        public void DataReader_BorrowedSdkEmptySheetIgnoresExtensionOwnedCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Empty");
            document.Save();
            Worksheet worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var payload = new OpenXmlUnknownElement("e", "metadata", "urn:officeimo:cell-ownership");
            payload.Append(CreateSdkOwnershipCell("number"));
            worksheet.Append(new WorksheetExtensionList(new WorksheetExtension(payload) {
                Uri = "{564B230A-5555-4000-8000-000000000001}"
            }));

            using DbDataReader reader = document.CreateDataReader();
            Assert.Equal(0, reader.FieldCount);
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData("shared-string")]
        [InlineData("style")]
        [InlineData("formula")]
        public void DataReader_BorrowedSdkValidatesOwnedCellReferences(string kind) {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Number");
            sheet.CellValue(2, 1, 1);
            document.Save();
            Worksheet worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Last().Append(CreateSdkOwnershipCell(kind));

            Action open = () => {
                using var reader = document.CreateDataReader(new ExcelReadOptions { UseCachedFormulaResult = false });
            };
            if (kind == "formula") Assert.Throws<NotSupportedException>(open);
            else Assert.Throws<InvalidDataException>(open);
        }

        private static Cell CreateSdkOwnershipCell(string kind) {
            var cell = new Cell { CellReference = "B2", DataType = CellValues.Number, CellValue = new CellValue("999") };
            if (kind == "shared-string") cell.DataType = CellValues.SharedString;
            else if (kind == "style") cell.StyleIndex = 999U;
            else if (kind == "formula") {
                cell.CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 999U };
                cell.CellValue = null;
            }
            return cell;
        }
    }
}
