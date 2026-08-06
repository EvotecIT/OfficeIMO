using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_QueryInspection_BoundsUnloadedNativeConnectionMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Data");
            ConnectionsPart connectionsPart = document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            byte[] payload = Encoding.UTF8.GetBytes(
                "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + new string('x', ExcelDocument.MaximumWorkbookConnectionMetadataCharacters)
                + "</connections>");
            using (var stream = new MemoryStream(payload, writable: false)) {
                connectionsPart.FeedData(stream);
            }

            Assert.False(connectionsPart.IsRootElementLoaded);
            Assert.Empty(document.GetQueryBackedTables());
            Assert.False(connectionsPart.IsRootElementLoaded);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_RangeTransfer_RenumbersSourceInCellImageBeforeSnapshot(bool move) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Discarded destination");
            sheet.SetInCellImage(1, 2, TinyPng, altText: "Transferred source");

            if (move) {
                sheet.MoveRange("B1", "A1");
            } else {
                sheet.CopyRange("B1", "A1");
            }

            ExcelInCellImage[] images = sheet.GetInCellImages()
                .OrderBy(image => image.CellReference, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(move ? new[] { "A1" } : new[] { "A1", "B1" }, images.Select(image => image.CellReference));
            Assert.All(images, image => {
                Assert.Equal("Transferred source", image.AltText);
                Assert.Equal(TinyPng, image.Bytes);
            });
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_TableAndQuerySchema_RejectInvalidXmlHeadersBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Original");
            sheet.CellValue(2, 1, "Keep");
            sheet.AddTable("A1:A2", true, "DataTable", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            string originalXml = table.OuterXml;

            ArgumentException tableException = Assert.Throws<ArgumentException>(() =>
                sheet.SetTableSchema("DataTable", new[] { "Bad\u0001Header" }));
            Assert.Equal("columnNames", tableException.ParamName);
            Assert.Equal(originalXml, table.OuterXml);
            Assert.Equal("Original", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());

            int partCount = document.WorkbookPartRoot.Parts.Count();
            ArgumentException queryException = Assert.Throws<ArgumentException>(() =>
                document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                    ConnectionName = "InvalidXml",
                    WorksheetName = sheet.Name,
                    TableName = "InvalidXmlResults",
                    ColumnNames = new[] { "Bad\u0001Header" }
                }));
            Assert.Equal("columnNames", queryException.ParamName);
            Assert.Equal(partCount, document.WorkbookPartRoot.Parts.Count());
            Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
        }
    }
}
