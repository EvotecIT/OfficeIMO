using System;
using System.IO;
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
                Assert.Contains("Unsupported formulas", calculationDiagnostics);
                Assert.Contains("Formula dependency issues", calculationDiagnostics);
                Assert.Contains("Unsupported formulas", Assert.Throws<InvalidOperationException>(() =>
                    report.EnsureCan(ExcelPreflightCapability.CalculateFormulas)).Message);

                string cachedValueDiagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.Contains("Missing formula caches", cachedValueDiagnostics);
                Assert.Contains("Formula dependency issues", cachedValueDiagnostics);
            }
        }
    }
}
