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
            sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15D)));
            X.SheetFormatProperties format = sheet.WorksheetPart.Worksheet.GetFirstChild<X.SheetFormatProperties>()!;
            format.BaseColumnWidth = baseWidth;
            format.DefaultColumnWidth = storedWidth.HasValue ? new DocumentFormat.OpenXml.DoubleValue(storedWidth.Value) : null;

            ExcelImage image = sheet.AddImageToRange("A1:C2", OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue)));
            Assert.Equal(expectedPixels, image.Width.ToPixels());
            Assert.Equal((double)expectedPixels, sheet.Range("A1:C2").CreateVisualSnapshot().Columns.Sum(column => column.Width));
            byte[] package = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx });
            using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(package));
            Assert.Equal(expectedPixels, reopened.Sheets[0].Images.Single().Width.ToPixels());
            Assert.Equal((double)expectedPixels, reopened.Sheets[0].Range("A1:C2").CreateVisualSnapshot().Columns.Sum(column => column.Width));
        }

        [Fact]
        public void ExcelChart_TwoCellAnchorUsesImplicitWidthAndRefreshesWhenBaseChanges() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("ImplicitChart");
            SetMissingGridFont(document);
            sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15D)));
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

            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo before));
            Assert.Equal(192, before.WidthPixels);
            format.BaseColumnWidth = 12U;
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo after));
            Assert.Equal(288, after.WidthPixels);
        }

        [Theory]
        [InlineData(null)]
        [InlineData(false)]
        [InlineData(true)]
        public void ExcelDrawing_StoredColumnWidthAppliesRegardlessOfCustomWidthFlag(bool? customWidth) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("StoredWidth");
            SetMissingGridFont(document);
            sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15D)));
            sheet.WorksheetPart.Worksheet.GetFirstChild<X.SheetFormatProperties>()!.BaseColumnWidth = 8U;
            sheet.WorksheetPart.Worksheet.AddChild(new X.Columns(new X.Column {
                Min = 1U, Max = 1U, Width = 20D,
                CustomWidth = customWidth.HasValue ? new DocumentFormat.OpenXml.BooleanValue(customWidth.Value) : null
            }), true);
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelImage image = sheet.AddImageToRange("A1:A2", OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue)));
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("2"), new Xdr.RowOffset("0")));

            Assert.Equal(140, image.Width.ToPixels());
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo before));
            Assert.Equal(140, before.WidthPixels);
            Assert.Equal(140, sheet.Range("A1:A2").ExportImage(OfficeImageExportFormat.Png).Width);
            using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx })));
            Assert.Equal(140, reopened.Sheets[0].Images.Single().Width.ToPixels());
            Assert.True(reopened.Sheets[0].Charts.Single().TryGetSnapshot(out ExcelChartInfo after));
            Assert.Equal(140, after.WidthPixels);
            Assert.Equal(140, reopened.Sheets[0].Range("A1:A2").ExportImage(OfficeImageExportFormat.Png).Width);
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
