using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelRange_ImageExportKeepsAnchorsInsideTallRequestedRanges() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("TallRange");
            sheet.AddImage(
                15000,
                1,
                CreateSolidPng(8, 8, OfficeColor.FromRgb(37, 99, 235)),
                "image/png",
                widthPixels: 8,
                heightPixels: 8,
                name: "DeepAnchor");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A20000").CreateVisualSnapshot(
                new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal("TallRange!DeepAnchor", image.Source);
        }

        [Fact]
        public void ExcelChart_ImageExportSkipsWorksheetGeometryForOneCellAnchors() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("OneCell");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            X.Row invalidRow = new X.Row();
            invalidRow.SetAttribute(new OpenXmlAttribute("r", string.Empty, "invalid"));
            sheet.WorksheetPart.Worksheet!.GetFirstChild<X.SheetData>()!.Append(invalidRow);

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.True(snapshot.WidthPixels > 0);
            Assert.True(snapshot.HeightPixels > 0);
        }

        [Fact]
        public void ExcelChart_ImageExportBoundsDuplicateRowDefinitionWork() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("DuplicateRows");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(
                document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("1"), new Xdr.RowOffset("0")));

            X.SheetData sheetData = sheet.WorksheetPart.Worksheet!.GetFirstChild<X.SheetData>()!;
            X.Row duplicate = new X.Row { RowIndex = 1U };
            sheetData.RemoveAllChildren<X.Row>();
            for (int index = 0; index < 100_000; index++) {
                sheetData.Append((X.Row)duplicate.CloneNode(true));
            }
            X.Row invalidRow = new X.Row();
            invalidRow.SetAttribute(new OpenXmlAttribute("r", string.Empty, "invalid"));
            sheetData.Append(invalidRow);

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.True(snapshot.WidthPixels > 0);
            Assert.True(snapshot.HeightPixels > 0);
        }
    }
}
