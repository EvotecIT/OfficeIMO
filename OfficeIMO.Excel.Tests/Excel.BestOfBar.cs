using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelBestOfBar_SafePreflight_DoesNotForceFullRecalculationForStructuredReferenceFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.SafePreflightFormula.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 10);
                sheet.CellValue(3, 1, 20);
                sheet.AddTable("A1:A3", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);
                sheet.CellFormula(2, 3, "SUM(Sales[Amount])");
                document.Save(filePath, openExcel: false, options: new ExcelSaveOptions { SafePreflight = true });
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            CalculationProperties? calculationProperties = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<CalculationProperties>();

            Assert.Null(calculationProperties?.FullCalculationOnLoad);
            Assert.Null(calculationProperties?.ForceFullCalculation);
        }

        [Fact]
        public void Test_ExcelBestOfBar_ExplicitFullCalculationRequest_StillWritesCalculationProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.FullCalcExplicit.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 10);
                sheet.CellFormula(1, 2, "A1*2");
                document.Save(filePath, openExcel: false, options: new ExcelSaveOptions { ForceFullCalculationOnOpen = true });
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            CalculationProperties calculationProperties = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<CalculationProperties>()!;

            Assert.True(calculationProperties.FullCalculationOnLoad!.Value);
            Assert.True(calculationProperties.ForceFullCalculation!.Value);
        }

        [Fact]
        public void Test_ExcelBestOfBar_CategoryAxisScale_WritesLineChartCategoryAxisUnits() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.CategoryAxisScale.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Trend");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 12);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 18);

                sheet.Chart("A1:B4")
                    .Line()
                    .Title("Trend")
                    .At(1, 5)
                    .SetCategoryAxisScale(majorUnit: 2, minorUnit: 1, reverseOrder: true);

                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorksheetPart worksheetPart = GetWorksheetPartWithCharts(spreadsheet);
            C.DateAxis categoryAxis = worksheetPart.DrawingsPart!.ChartParts
                .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.DateAxis>())
                .First(axis => axis != null)!;

            Assert.Equal(2D, categoryAxis.GetFirstChild<C.MajorUnit>()!.Val!.Value);
            Assert.Equal(1D, categoryAxis.GetFirstChild<C.MinorUnit>()!.Val!.Value);
            Assert.Equal(C.OrientationValues.MaxMin, categoryAxis.GetFirstChild<C.Scaling>()!.GetFirstChild<C.Orientation>()!.Val!.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(spreadsheet).ToList();
            Assert.True(errors.Count == 0, FormatValidationErrors(errors));
        }

        [Fact]
        public void Test_ExcelBestOfBar_TableStyleCatalog_FlagsCrossHostRecommendations() {
            ExcelTableStyleCompatibilityInfo stable = ExcelTableStyleCatalog.Analyze(ExcelTableStyle.TableStyleMedium2);
            ExcelTableStyleCompatibilityInfo heavier = ExcelTableStyleCatalog.Analyze(ExcelTableStyle.TableStyleDark11);
            ExcelTableStyleCompatibilityInfo unknown = ExcelTableStyleCatalog.Analyze("CustomStyle");

            Assert.True(stable.IsRecommended);
            Assert.False(heavier.IsRecommended);
            Assert.Contains("cross-host", heavier.Warning, StringComparison.OrdinalIgnoreCase);
            Assert.False(unknown.IsBuiltIn);
            Assert.Contains(nameof(ExcelTableStyle.TableStyleMedium2), ExcelTableStyleCatalog.GetRecommendedNames());
        }

        [Fact]
        public void Test_ExcelBestOfBar_WriteReservation_RoundTripsSeparatelyFromWorkbookProtection() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.WriteReservation.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
                    ReadOnlyRecommended = true,
                    UserName = "Reviewer",
                    LegacyPasswordHash = "CAFE"
                });
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions { LegacyPasswordHash = "BEEF" });
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookWriteReservationInfo reservation = document.GetWriteReservation();

                Assert.True(reservation.Exists);
                Assert.True(reservation.ReadOnlyRecommended);
                Assert.Equal("Reviewer", reservation.UserName);
                Assert.Equal("CAFE", reservation.LegacyPasswordHash);
                Assert.True(document.IsWorkbookProtected);
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
            var children = workbook.ChildElements.ToList();
            int fileSharingIndex = children.FindIndex(element => element is FileSharing);
            int workbookPropertiesIndex = children.FindIndex(element => element is WorkbookProperties);
            int workbookProtectionIndex = children.FindIndex(element => element is WorkbookProtection);

            Assert.InRange(fileSharingIndex, 0, int.MaxValue);
            Assert.InRange(workbookProtectionIndex, 0, int.MaxValue);
            if (workbookPropertiesIndex >= 0) {
                Assert.True(fileSharingIndex < workbookPropertiesIndex);
            }
            Assert.True(fileSharingIndex < workbookProtectionIndex);
        }

        [Fact]
        public void Test_ExcelBestOfBar_RuntimePreflight_ReportsCurrentCultureAndWarnings() {
            ExcelRuntimePreflightReport report = ExcelRuntimePreflight.InspectCurrent();

            Assert.False(string.IsNullOrWhiteSpace(report.FrameworkDescription));
            Assert.False(string.IsNullOrWhiteSpace(report.OSDescription));
            Assert.NotNull(report.CurrentCultureName);
            Assert.NotNull(report.CurrentUICultureName);
            Assert.NotNull(report.Warnings);
        }

        [Fact]
        public void Test_ExcelBestOfBar_RowHeight_DoesNotSpillIntoNeighborRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.RowHeight.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Rows");
                sheet.CellValue(3, 1, "Tall");
                sheet.CellValue(4, 1, "Normal");
                sheet.SetRowHeight(3, 42D);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Row row3 = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Row>().First(row => row.RowIndex!.Value == 3U);
            Row row4 = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Row>().First(row => row.RowIndex!.Value == 4U);

            Assert.Equal(42D, row3.Height!.Value);
            Assert.True(row3.CustomHeight!.Value);
            Assert.Null(row4.Height);
            Assert.Null(row4.CustomHeight);
        }

        [Fact]
        public void Test_ExcelBestOfBar_DirectDataSetLargeAppend_SaveRemainsReadableAndValid() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.LargeDataSet.xlsx");
            var dataSet = new DataSet("LargeAppend");
            var table = new DataTable("Rows");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            for (int i = 1; i <= 2500; i++) {
                table.Rows.Add(i, "Item " + i.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            dataSet.Tables.Add(table);

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet, createTables: true, autoFit: false);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Id", document.Sheets[0].CellAt(1, 1).GetValue<string>());
                Assert.Equal(2500D, document.Sheets[0].CellAt(2501, 1).GetValue<double>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelBestOfBar_NumberFormatCatalog_ExposesPresetNames() {
            Assert.Contains(nameof(ExcelNumberPreset.Currency), ExcelNumberFormats.GetPresetNames());
            Assert.Equal("0.00%", ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals: 2));
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_AppliesHeaderFormats() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.xlsx");

            var table = new DataTable("Sales");
            table.Columns.Add("Id", typeof(string));
            table.Columns.Add("Revenue", typeof(decimal));
            table.Columns.Add("Rate", typeof(decimal));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add("00042", 1234.5M, 0.125M, new DateTime(2026, 6, 23));

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, 1, 1, includeHeaders: true, tableName: "Sales", style: ExcelTableStyle.TableStyleMedium2);

                var plan = new ExcelColumnFormatPlan()
                    .Text("Id")
                    .Currency(0, System.Globalization.CultureInfo.GetCultureInfo("en-US"), "Revenue")
                    .Add("Rate", ExcelNumberPreset.Percent, decimals: 1)
                    .Add("Created", ExcelNumberPreset.DateShort);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(plan, autoFit: true);
                IReadOnlyList<ExcelColumnFormatResult> missing = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Add("Missing", ExcelNumberPreset.Integer));

                Assert.All(results, result => Assert.True(result.Applied));
                Assert.Equal(new[] { 1, 2, 3, 4 }, results.Select(result => result.ColumnIndex!.Value).ToArray());
                Assert.Single(missing);
                Assert.False(missing[0].Applied);
                Assert.Contains("Missing", missing[0].Warning);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("@", GetCellNumberFormatCode(spreadsheet, cells["A2"]));
            Assert.Contains("$", GetCellNumberFormatCode(spreadsheet, cells["B2"]) ?? string.Empty);
            Assert.Equal("0.0%", GetCellNumberFormatCode(spreadsheet, cells["C2"]));
            Assert.Equal("yyyy-mm-dd", GetCellNumberFormatCode(spreadsheet, cells["D2"]));
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_HonorsExactAndExplicitHeaderRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.HeaderRows.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Report");
                sheet.CellValue(2, 1, "Value");
                sheet.CellValue(2, 2, "  Value  ");
                sheet.CellValue(3, 1, 10);
                sheet.CellValue(3, 2, 20);

                var exactPlan = new ExcelColumnFormatPlan().Add("  Value  ", ExcelNumberPreset.Integer);
                IReadOnlyList<ExcelColumnFormatResult> exactResults = sheet.ApplyColumnFormatPlan(
                    exactPlan,
                    headerRow: 2,
                    options: new ExcelReadOptions { NormalizeHeaders = false });

                var titlePlan = new ExcelColumnFormatPlan().Currency(System.Globalization.CultureInfo.GetCultureInfo("en-US"), "Value");
                IReadOnlyList<ExcelColumnFormatResult> titleResults = sheet.ApplyColumnFormatPlan(
                    titlePlan,
                    headerRow: 2);

                Assert.Single(exactResults);
                Assert.True(exactResults[0].Applied);
                Assert.Equal(2, exactResults[0].ColumnIndex);
                Assert.Single(titleResults);
                Assert.True(titleResults[0].Applied);
                Assert.Equal(1, titleResults[0].ColumnIndex);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Contains("$", GetCellNumberFormatCode(spreadsheet, cells["A3"]) ?? string.Empty);
            Assert.Equal("#,##0", GetCellNumberFormatCode(spreadsheet, cells["B3"]));
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_RejectsBlankExplicitHeaderRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.BlankExplicitHeader.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(3, 1, 42);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Add("Column1", ExcelNumberPreset.Integer),
                    headerRow: 2);

                Assert.Single(results);
                Assert.False(results[0].Applied);
                Assert.Null(results[0].ColumnIndex);
                document.Save();
            }
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_NormalizesDefaultHeaderLookup() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.DefaultHeaderLookup.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "  Value  ");
                sheet.CellValue(2, 1, 42);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Add("  Value  ", ExcelNumberPreset.Integer));

                Assert.Single(results);
                Assert.True(results[0].Applied);
                Assert.Equal(1, results[0].ColumnIndex);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("#,##0", GetCellNumberFormatCode(spreadsheet, cells["A2"]));
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_AllowsHeaderOnlyDirectExports() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.HeaderOnlyDirect.xlsx");

            var table = new DataTable("Counts");
            table.Columns.Add("Count", typeof(int));

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, 1, 1, includeHeaders: true, tableName: "Counts", style: ExcelTableStyle.TableStyleMedium2);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Add("Count", ExcelNumberPreset.Integer));

                Assert.Single(results);
                Assert.True(results[0].Applied);
                Assert.Equal(1, results[0].ColumnIndex);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Contains("A1", cells.Keys);
            Assert.DoesNotContain("A2", cells.Keys);
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_FormatsPreservedOverlayHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.OverlayHeader.xlsx");

            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, 1, 1, includeHeaders: true, tableName: "Rows", style: ExcelTableStyle.TableStyleMedium2);
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 2, 42);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Add("Score", ExcelNumberPreset.Integer));

                Assert.Single(results);
                Assert.True(results[0].Applied);
                Assert.Equal(2, results[0].ColumnIndex);
                document.Save(filePath, openExcel: false, options: new ExcelSaveOptions { DisableFastPackageWriter = true });
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("#,##0", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_UpdatesMaterializedDomForStandardSaveFallback() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.MaterializedFallback.xlsx");

            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Amount", typeof(decimal));
            table.Rows.Add("Alpha", 12.5M);

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, 1, 1, includeHeaders: true, tableName: "Rows", style: ExcelTableStyle.TableStyleMedium2);
                Assert.True(sheet.TryGetCellText(2, 1, out _));

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan().Currency(System.Globalization.CultureInfo.GetCultureInfo("en-US"), "Amount"));

                Assert.Single(results);
                Assert.True(results[0].Applied);
                Assert.Equal(2, results[0].ColumnIndex);
                document.Save(filePath, openExcel: false, options: new ExcelSaveOptions { DisableFastPackageWriter = true });
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Contains("$", GetCellNumberFormatCode(spreadsheet, cells["B2"]) ?? string.Empty);
        }

        [Fact]
        public void Test_ExcelBestOfBar_ColumnFormatPlan_AutoFitKeepsDirectPackageWriter() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelBestOfBar.ColumnFormatPlan.DirectAutoFit.xlsx");

            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Amount", typeof(decimal));
            table.Rows.Add("Long descriptive customer name", 1234.5M);

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, 1, 1, includeHeaders: true, tableName: "Rows", style: ExcelTableStyle.TableStyleMedium2);

                IReadOnlyList<ExcelColumnFormatResult> results = sheet.ApplyColumnFormatPlan(
                    new ExcelColumnFormatPlan()
                        .Text("Name")
                        .Currency(System.Globalization.CultureInfo.GetCultureInfo("en-US"), "Amount"),
                    autoFit: true);

                Assert.All(results, result => Assert.True(result.Applied));
                document.Save();
                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("@", GetCellNumberFormatCode(spreadsheet, cells["A2"]));
            Assert.Contains("$", GetCellNumberFormatCode(spreadsheet, cells["B2"]) ?? string.Empty);
        }
    }
}
