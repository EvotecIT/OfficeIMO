using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly byte[] TinyPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

        [Fact]
        public void Test_InCellImage_RoundTripsAndFollowsStructuralEdits() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.SetInCellImage(2, 2, TinyPng, altText: "Status badge");

                sheet.InsertRows(2);
                sheet.InsertColumns(2);

                ExcelInCellImage image = Assert.Single(sheet.GetInCellImages());
                Assert.Equal("C3", image.CellReference);
                Assert.Equal("Status badge", image.AltText);
                Assert.Equal(TinyPng, image.Bytes);
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            stream.Position = 0;
            using var loaded = ExcelDocument.Load(stream);
            ExcelInCellImage roundTripped = Assert.Single(loaded.Sheets[0].GetInCellImages());
            Assert.Equal("C3", roundTripped.CellReference);
            Assert.Equal("image/png", roundTripped.ContentType);
            Assert.Equal(TinyPng, roundTripped.Bytes);
            Assert.Empty(loaded.ValidateOpenXml());
        }

        [Fact]
        public void Test_InCellImage_CopyMoveAndRemovePreserveNativeValue() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Logo");

            sheet.Range("A1").CopyTo("B2");
            Assert.Equal(new[] { "A1", "B2" }, sheet.GetInCellImages().Select(image => image.CellReference));

            sheet.Range("B2").MoveTo("C3");
            Assert.Equal(new[] { "A1", "C3" }, sheet.GetInCellImages().Select(image => image.CellReference));
            Assert.True(sheet.RemoveInCellImage(3, 3));
            Assert.False(sheet.RemoveInCellImage(3, 3));
            Assert.Equal("A1", Assert.Single(sheet.GetInCellImages()).CellReference);
        }

        [Fact]
        public void Test_InCellImage_FollowsSortAndSurvivesFilterAndResizeMetadata() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.CellValue(1, 1, "Key");
            sheet.CellValue(1, 2, "Image");
            sheet.CellValue(2, 1, "Zulu");
            sheet.CellValue(3, 1, "Alpha");
            sheet.SetInCellImage(2, 2, TinyPng, altText: "Sorted badge");
            sheet.AddAutoFilter("A1:B3");
            sheet.SetColumnWidth(2, 24D);

            sheet.Range("A1:B3").SortByColumn(1, ascending: true, hasHeader: true);

            ExcelInCellImage image = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("B3", image.CellReference);
            Assert.Equal("Sorted badge", image.AltText);
            Assert.NotEmpty(sheet.GetAutoFilters());
        }

        [Fact]
        public void Test_InCellImage_ReplacesAndRemovesInlineStringPayload() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.CellValue(1, 1, "Old value");
            Cell cell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>());
            cell.CellValue = null;
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text("Old value"));

            sheet.SetInCellImage(1, 1, TinyPng, altText: "Replacement");

            Assert.Null(cell.InlineString);
            Assert.Equal("#VALUE!", cell.CellValue!.Text);
            Assert.Single(sheet.GetInCellImages());

            Assert.True(sheet.RemoveInCellImage(1, 1));
            Assert.Null(cell.InlineString);
            Assert.Null(cell.CellValue);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
