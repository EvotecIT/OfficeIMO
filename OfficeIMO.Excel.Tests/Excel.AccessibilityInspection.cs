using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelAccessibilityInspectionTests {
        [Fact]
        public void WellFormedWorkbookPassesTheStrictProfile() {
            using ExcelDocument document = ExcelDocument.Create();
            document.Properties.Title = "Quarterly budget";
            ExcelSheet sheet = document.AddSheet("Budget");
            WriteBlock(sheet, firstRow: 1, rows: 3, columns: 2);
            sheet.AddTable("A1:B3", hasHeader: true, name: "BudgetTable");
            sheet.AddImage(5, 1, Png(), new ExcelImageOptions { ContentType = "image/png" }).SetAltText("Company logo");
            sheet.AddImage(5, 3, Png(), new ExcelImageOptions { ContentType = "image/png" }).Decorative();
            sheet.SetHyperlink("C1", "https://officeimo.com", "OfficeIMO documentation");

            OfficeAccessibilityReport report = document.InspectAccessibility(ExcelAccessibilityOptions.ForProfile(OfficeAccessibilityPolicyProfile.Strict));

            Assert.Equal("Excel", report.Format);
            Assert.Empty(report.Findings);
        }

        [Fact]
        public void PicturesAndChartsNeedAlternativeTextUnlessDecorative() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Value");
            ExcelImage picture = sheet.AddImage(2, 2, Png(), new ExcelImageOptions { ContentType = "image/png" });
            ExcelChart chart = sheet.AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B" },
                new[] { new OfficeSeries("Sales", new[] { 1d, 2d }) }), new ExcelChartOptions { Row = 6, Column = 4 });

            List<OfficeAccessibilityFinding> missing = Findings(document, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText);
            Assert.Equal(2, missing.Count);
            Assert.All(missing, finding => Assert.Equal(OfficeAccessibilitySeverity.Error, finding.Severity));
            Assert.Contains(missing, finding => finding.Location == "Data!B2" && finding.Message.Contains("picture") && finding.ObjectName == picture.Name);
            Assert.Contains(missing, finding => finding.Location == "Data!D6" && finding.Message.Contains("chart") && finding.LocationIndex == 0);

            picture.Decorative();
            chart.SetAltText("Sales rose from 1 to 2.");
            Assert.True(picture.IsDecorative);
            Assert.Equal("Sales rose from 1 to 2.", chart.AltText);
            Assert.Empty(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText));

            chart.Decorative();
            Assert.True(chart.IsDecorative);
            Assert.Equal(string.Empty, chart.AltText);
            Assert.Empty(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText));
        }

        [Fact]
        public void DecorativeFlagSurvivesSaveAndReload() {
            using var stream = new MemoryStream();
            using (ExcelDocument document = ExcelDocument.Create(stream)) {
                document.AddSheet("Data").AddImage(1, 1, Png(), new ExcelImageOptions { ContentType = "image/png" }).Decorative();
                document.Save();
            }
            stream.Position = 0;
            using ExcelDocument reloaded = ExcelDocument.Load(stream);
            Assert.True(reloaded.Sheets[0].Images.Single().IsDecorative);
            Assert.Empty(Findings(reloaded, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText));
        }

        [Fact]
        public void TablesWithoutHeaderRowAreErrors() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            WriteBlock(sheet, firstRow: 1, rows: 3, columns: 2);
            WriteBlock(sheet, firstRow: 1, rows: 3, columns: 2, firstColumn: 4);
            sheet.AddTable("A1:B3", hasHeader: false, name: "NoHeader");
            sheet.AddTable("D1:E3", hasHeader: true, name: "WithHeader");

            OfficeAccessibilityFinding finding = Assert.Single(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingTableHeader));
            Assert.Equal(OfficeAccessibilitySeverity.Error, finding.Severity);
            Assert.Equal("Data!A1:B3", finding.Location);
            Assert.Equal("NoHeader", finding.ObjectName);
        }

        [Fact]
        public void MergedCellsAreStructuralWarningsAndStrictErrors() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Heading");
            sheet.Range("A1:C1").Merge();

            OfficeAccessibilityFinding finding = Assert.Single(Findings(document, OfficeDiagnosticCodes.AccessibilityMergedCells));
            Assert.Equal(OfficeAccessibilitySeverity.Warning, finding.Severity);
            Assert.Equal("Data!A1:C1", finding.Location);
            Assert.Equal(OfficeAccessibilitySeverity.Error,
                Assert.Single(Findings(document, OfficeDiagnosticCodes.AccessibilityMergedCells, OfficeAccessibilityPolicyProfile.Strict)).Severity);
            Assert.Empty(Findings(document, OfficeDiagnosticCodes.AccessibilityMergedCells, options: new ExcelAccessibilityOptions { CheckMergedCells = false }));
        }

        [Fact]
        public void DefaultSheetNamesAreReportedOnlyForSheetsWithContent() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Sheet1").Cell(1, 1).SetValue("Data");
            document.AddSheet("Arkusz2").Cell(1, 1).SetValue("Dane");
            document.AddSheet("Sheet3");
            document.AddSheet("Sheet names").Cell(1, 1).SetValue("Named");

            List<OfficeAccessibilityFinding> findings = Findings(document, OfficeDiagnosticCodes.AccessibilityDefaultSheetName);
            Assert.Equal(new[] { "Sheet1", "Arkusz2" }, findings.Select(finding => finding.Location));
            Assert.Equal(new int?[] { 0, 1 }, findings.Select(finding => finding.LocationIndex));
        }

        [Fact]
        public void HiddenSheetsWithContentAreReported() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Visible").Cell(1, 1).SetValue("Shown");
            ExcelSheet hidden = document.AddSheet("Lookups");
            hidden.Cell(1, 1).SetValue("Secret");
            hidden.SetHidden(true);
            ExcelSheet veryHidden = document.AddSheet("Calc");
            veryHidden.Cell(1, 1).SetValue(42);
            veryHidden.SetVeryHidden(true);
            document.AddSheet("Empty").SetHidden(true);

            List<OfficeAccessibilityFinding> findings = Findings(document, OfficeDiagnosticCodes.AccessibilityHiddenSheetContent);
            Assert.Equal(new[] { "Lookups", "Calc" }, findings.Select(finding => finding.Location));
            Assert.Contains("very hidden", findings[1].Message);
        }

        [Fact]
        public void HyperlinkCellsNeedMeaningfulText() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Links");
            sheet.SetHyperlink("A1", "https://officeimo.com", "click here");
            sheet.SetHyperlink("A2", "https://officeimo.com/docs", "https://officeimo.com/docs");
            sheet.SetHyperlink("A3", "https://officeimo.com/guide", "Excel accessibility guide");

            List<OfficeAccessibilityFinding> findings = Findings(document, OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel);
            Assert.Equal(new[] { "Links!A1", "Links!A2" }, findings.Select(finding => finding.Location));
            Assert.All(findings, finding => Assert.Equal(OfficeAccessibilitySeverity.Error, finding.Severity));
        }

        [Fact]
        public void LowContrastCellsAreGroupedByFormatting() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Colors");
            for (int row = 1; row <= 3; row++) {
                sheet.Cell(row, 1).SetValue("Pale").SetColor(OfficeColor.FromRgb(0xCC, 0xCC, 0xCC)).SetFillColor(OfficeColor.White);
            }
            sheet.Cell(1, 2).SetValue("Readable").SetColor(OfficeColor.FromRgb(0x33, 0x33, 0x33)).SetFillColor(OfficeColor.White);
            // #949494 on white is about 3.03:1: too low for normal text, enough for large bold text.
            sheet.Cell(1, 3).SetValue("Large").SetColor(OfficeColor.FromRgb(0x94, 0x94, 0x94)).SetBold().SetFontSize(OfficeLength.Points(14));
            sheet.Cell(2, 3).SetValue("Small").SetColor(OfficeColor.FromRgb(0x94, 0x94, 0x94));

            List<OfficeAccessibilityFinding> findings = Findings(document, OfficeDiagnosticCodes.AccessibilityLowContrast);
            Assert.Equal(2, findings.Count);
            OfficeAccessibilityFinding pale = Assert.Single(findings, finding => finding.Location == "Colors!A1");
            Assert.StartsWith("3 cells", pale.Message);
            Assert.Equal(4.5D, pale.RequiredValue);
            Assert.True(pale.MeasuredValue < 2D);
            OfficeAccessibilityFinding small = Assert.Single(findings, finding => finding.Location == "Colors!C2");
            Assert.Equal(4.5D, small.RequiredValue);
        }

        [Fact]
        public void BlankRowsAndColumnsInsideDataAreHeuristicWarnings() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet spaced = document.AddSheet("Spaced");
            WriteBlock(spaced, firstRow: 1, rows: 2, columns: 2);
            WriteBlock(spaced, firstRow: 5, rows: 2, columns: 2);
            spaced.Cell(1, 4).SetValue("Notes");
            ExcelSheet compact = document.AddSheet("Compact");
            WriteBlock(compact, firstRow: 1, rows: 4, columns: 3);

            List<OfficeAccessibilityFinding> findings = Findings(document, OfficeDiagnosticCodes.AccessibilityBlankCellLayout);
            Assert.Equal(new[] { "Spaced!A3:D4", "Spaced!C1:C6" }, findings.Select(finding => finding.Location));
            Assert.All(findings, finding => {
                Assert.True(finding.IsHeuristic);
                Assert.Equal(OfficeAccessibilitySeverity.Warning, finding.Severity);
            });
            Assert.StartsWith("Rows 3-4", findings[0].Message);
            Assert.StartsWith("Column C", findings[1].Message);
            Assert.All(Findings(document, OfficeDiagnosticCodes.AccessibilityBlankCellLayout, OfficeAccessibilityPolicyProfile.Strict),
                finding => Assert.Equal(OfficeAccessibilitySeverity.Warning, finding.Severity));
        }

        [Fact]
        public void StrictProfileRequiresWorkbookTitle() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Data").Cell(1, 1).SetValue("Value");

            Assert.Empty(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle));
            OfficeAccessibilityFinding finding = Assert.Single(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle, OfficeAccessibilityPolicyProfile.Strict));
            Assert.Equal("Workbook", finding.Location);
            document.Properties.Title = "Data";
            Assert.Empty(Findings(document, OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle, OfficeAccessibilityPolicyProfile.Strict));
        }

        [Fact]
        public void QuotedSheetNamesAppearInLocations() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Q1 Sales");
            sheet.Cell(1, 1).SetValue("Total");
            sheet.Range("A1:B1").Merge();

            Assert.Equal("'Q1 Sales'!A1:B1", Assert.Single(Findings(document, OfficeDiagnosticCodes.AccessibilityMergedCells)).Location);
        }

#pragma warning disable CS0618 // The obsolete report forwards to InspectAccessibility.
        [Fact]
        public void ObsoleteAnalyzeAccessibilityKeepsItsCategories() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Heading");
            sheet.Range("A1:B1").Merge();
            sheet.AddImage(3, 1, Png(), new ExcelImageOptions { ContentType = "image/png" });

            ExcelAccessibilityReport report = document.AnalyzeAccessibility();

            Assert.Contains(report.Findings, finding => finding.Category == "MergedCells" && finding.SheetName == "Data" && finding.Address == "A1:B1");
            Assert.Contains(report.Findings, finding => finding.Category == "ImageAltText" && finding.Severity == ExcelFindingSeverity.Error && finding.Address == "A3");
        }
#pragma warning restore CS0618

        private static List<OfficeAccessibilityFinding> Findings(ExcelDocument document, string code,
            OfficeAccessibilityPolicyProfile profile = OfficeAccessibilityPolicyProfile.Default, ExcelAccessibilityOptions? options = null) =>
            document.InspectAccessibility(options ?? ExcelAccessibilityOptions.ForProfile(profile)).Findings
                .Where(finding => finding.Code == code).ToList();

        private static void WriteBlock(ExcelSheet sheet, int firstRow, int rows, int columns, int firstColumn = 1) {
            for (int row = firstRow; row < firstRow + rows; row++) {
                for (int column = firstColumn; column < firstColumn + columns; column++) {
                    sheet.Cell(row, column).SetValue("R" + row + "C" + column);
                }
            }
        }

        private static byte[] Png() {
            var image = new OfficeRasterImage(2, 2, OfficeColor.Transparent);
            image.Fill(OfficeColor.Blue);
            return OfficePngWriter.Encode(image);
        }
    }
}
