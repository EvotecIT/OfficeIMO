using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_GoogleSheetsBatch_SizesGridForDimensionsAndProtectionExceptions() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsMetadataGridSizing.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.SetRowHeight(2000, 24);
                sheet.SetColumnWidth(28, 16);
                sheet.Protect();
                var options = new GoogleSheetsSaveOptions();
                options.Protection.UnprotectedRangesBySheet["Data"] = new List<string> { "A1:A2100" };

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document, options);

                GoogleSheetsAddSheetRequest addSheet = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddSheetRequest>(),
                    request => request.SheetName == "Data");
                Assert.Equal(2100, addSheet.RowCount);
                Assert.Equal(28, addSheet.ColumnCount);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_HonorsWorkbookDateSystemForValidationDates() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheets1904Validation.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                document.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.ValidationDate(
                    "A1:A2",
                    DataValidationOperatorValues.Between,
                    new DateTime(2024, 1, 1),
                    new DateTime(2024, 12, 31));

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsSetDataValidationRequest validation = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsSetDataValidationRequest>());
                Assert.Equal(new[] { "2024-01-01", "2024-12-31" }, validation.Rule.Values);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsCheckpoint_SeparatesSameNamedSheetScopedRanges() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsScopedNameHashes.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet north = document.AddWorksheet("North");
                ExcelSheet south = document.AddWorksheet("South");
                north.SetNamedRange("LocalData", "A1:A2", save: false);
                south.SetNamedRange("LocalData", "B1:B2", save: false);

                GoogleSheetsSyncCheckpoint checkpoint = GoogleSheetsDiffPlanner.CreateCheckpoint(document);

                Assert.Contains("name/sheet/North/LocalData", checkpoint.ContentHashes.Keys);
                Assert.Contains("name/sheet/South/LocalData", checkpoint.ContentHashes.Keys);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsPivot_UsesOnlyHeadersInsideSourceRange() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsPivotSourceHeaders.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 3, "Region");
                sheet.CellValue(1, 4, "Value");
                sheet.CellValue(2, 3, "North");
                sheet.CellValue(2, 4, 10d);
                sheet.CellValue(3, 3, "South");
                sheet.CellValue(3, 4, 20d);
                sheet.AddPivotTable(
                    "C1:D3",
                    "F1",
                    name: "ScopedHeaders",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Value", DataConsolidateFunctionValues.Sum, "Total") });

                GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();

                GoogleSheetsAddPivotTableRequest pivot = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>());
                Assert.Equal(0, Assert.Single(pivot.Rows).SourceColumnOffset);
                Assert.Equal(1, Assert.Single(pivot.Values).SourceColumnOffset);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
