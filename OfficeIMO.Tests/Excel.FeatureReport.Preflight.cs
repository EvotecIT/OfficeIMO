using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FeatureReport_Preflight_AllowsCleanReportWorkbookWorkflows() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Clean.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "Beta");
                sheet.CellValue(3, 2, 20d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "Scores", style: TableStyle.TableStyleMedium4);
                sheet.CellFormula(4, 2, "SUM(B2:B3)");
                Assert.Equal(1, document.Calculate());
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                Assert.True(report.Can(ExcelPreflightCapability.EditWorkbookStructure));
                Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.True(report.Can(ExcelPreflightCapability.BindTemplate));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Same(report, report.EnsureCan(ExcelPreflightCapability.ExportPdfReport));

                string markdown = report.ToMarkdown();
                Assert.Contains("## Capability Preflight", markdown);
                Assert.Contains("ExportPdfReport", markdown);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksStructureTemplateAndPdfWhenPreserveOnlyPartsExist() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.PreserveOnly.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.SetHyperlink(2, 1, "https://example.org/spec", display: "Spec");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                Assert.False(report.Can(ExcelPreflightCapability.EditWorkbookStructure));
                Assert.False(report.Can(ExcelPreflightCapability.BindTemplate));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                Assert.Contains("External workbook links", string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.EditWorkbookStructure)));
                Assert.Contains("https://example.org/spec", string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport)));

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                    report.EnsureCan(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("ExportPdfReport", exception.Message);
                Assert.Contains("External workbook links", exception.Message);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksFormulaCalculationAndCachedReadsWhenFormulaStateIsUnsafe() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Formulas.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Calc");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(2, 1, "B");
                sheet.CellFormula(1, 2, "UNIQUE(A1:A2)");
                sheet.CellFormula(2, 2, "B1+1");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Formula dependency issues");

                string calculationDiagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.CalculateFormulas));
                Assert.Contains("Formula calculation blockers", calculationDiagnostics);
                Assert.Contains("Formula dependency issues", calculationDiagnostics);
                Assert.Contains("Formula calculation blockers", Assert.Throws<InvalidOperationException>(() =>
                    report.EnsureCan(ExcelPreflightCapability.CalculateFormulas)).Message);

                string cachedValueDiagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.Contains("Missing formula caches", cachedValueDiagnostics);
                Assert.Contains("Formula dependency issues", cachedValueDiagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksCellEditsForSignedWorkbooks() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Signed.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Signed");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Ready");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.False(report.Can(ExcelPreflightCapability.EditCellValues));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Digital signatures");

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.EditCellValues));
                Assert.Contains("Digital signatures", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksCachedReadsWhenFormulaCachesAreDirty() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.DirtyFormulaCaches.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                Assert.Equal(1, document.Calculate());
                document.InvalidateFormulas();
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.Features, feature => feature.Name == "Dirty formula caches");

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.Contains("Dirty formula caches", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsSupportedFormulaChainsWithoutCachedResultsToCalculate() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.FormulaChain.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                sheet.CellFormula(1, 3, "B1+1");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Formula dependency issues");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.CalculateFormulas));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedChartTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.SurfaceChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Charts");
                sheet.CellValue(1, 1, "Zone");
                sheet.CellValue(1, 2, "Low");
                sheet.CellValue(1, 3, "High");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(2, 3, 15d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 12d);
                sheet.CellValue(3, 3, 18d);
                sheet.AddChartFromRange("A1:C3", row: 1, column: 5, widthPixels: 320, heightPixels: 180,
                    type: ExcelChartType.Surface, title: "Surface Chart");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported charts", diagnostics);
                Assert.Contains("Surface", diagnostics);
            }
        }

        private static void AddDigitalSignatureMetadata(string filePath) {
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo /></Signature>");

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            spreadsheet.AddDigitalSignatureOriginPart();
            DigitalSignatureOriginPart originPart = spreadsheet.DigitalSignatureOriginPart!;
            XmlSignaturePart signaturePart = originPart.AddNewPart<XmlSignaturePart>();
            using (var stream = new MemoryStream(signatureBytes)) {
                signaturePart.FeedData(stream);
            }

            ExtendedFilePropertiesPart appPart = spreadsheet.ExtendedFilePropertiesPart ?? spreadsheet.AddExtendedFilePropertiesPart();
            appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            appPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
            appPart.Properties.Save();
        }
    }
}
