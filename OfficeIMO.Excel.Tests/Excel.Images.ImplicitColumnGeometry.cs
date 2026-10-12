using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Theory]
        [InlineData(0U, null, 192)]
        [InlineData(1U, null, 48)]
        [InlineData(8U, null, 192)]
        [InlineData(9U, null, 216)]
        [InlineData(12U, null, 288)]
        [InlineData(15U, null, 336)]
        [InlineData(12U, 13D, 273)]
        public void ExcelImage_RangeUsesImplicitColumnGeometryBeforeAndAfterSave(uint baseWidth, double? storedWidth, int expectedPixels) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("ImplicitWidth");
            SetMissingGridFont(document);
            sheet.SetDefaultRowHeight(15D);
            X.SheetFormatProperties format = sheet.WorksheetPart.Worksheet.GetFirstChild<X.SheetFormatProperties>()!;
            format.BaseColumnWidth = baseWidth;
            format.DefaultColumnWidth = storedWidth.HasValue ? new DocumentFormat.OpenXml.DoubleValue(storedWidth.Value) : null;

            ExcelImage image = sheet.AddImageToRange("A1:C2", OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue)));
            Assert.Equal(expectedPixels, image.WidthPixels);
            byte[] package = document.ToBytes(ExcelFileFormat.Xlsx);
            using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(package));
            Assert.Equal(expectedPixels, reopened.Sheets[0].Images.Single().WidthPixels);
        }

        [Fact]
        public void ExcelChart_TwoCellAnchorUsesImplicitWidthAndRefreshesWhenBaseChanges() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("ImplicitChart");
            SetMissingGridFont(document);
            sheet.SetDefaultRowHeight(15D);
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            X.SheetFormatProperties format = sheet.WorksheetPart.Worksheet.GetFirstChild<X.SheetFormatProperties>()!;
            format.BaseColumnWidth = 8U;
            format.DefaultColumnWidth = null;
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("3"), new Xdr.ColumnOffset("0"), new Xdr.RowId("2"), new Xdr.RowOffset("0")));

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot before));
            Assert.Equal(192, before.WidthPixels);
            format.BaseColumnWidth = 12U;
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot after));
            Assert.Equal(288, after.WidthPixels);
        }

        private static void SetMissingGridFont(ExcelDocument document) {
            // This exercises the supported seven-pixel missing-font fallback on every host;
            // installed-face measurements are independently compared with native Excel.
            WorkbookStylesPart part = document.WorkbookPartRoot.WorkbookStylesPart
                ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
            part.Stylesheet ??= new X.Stylesheet(new X.Fonts(new X.Font(
                new X.FontName { Val = "OfficeIMO Missing Grid Font 28b06c8d" }, new X.FontSize { Val = 11D })),
                new X.Fills(new X.Fill()), new X.Borders(new X.Border()), new X.CellFormats(new X.CellFormat()));
            X.Font font = part.Stylesheet.Fonts!.Elements<X.Font>().First();
            font.GetFirstChild<X.FontName>()!.Val = "OfficeIMO Missing Grid Font 28b06c8d";
        }
    }
}
