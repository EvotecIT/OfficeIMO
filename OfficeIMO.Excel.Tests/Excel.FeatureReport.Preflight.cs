using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FeatureReport_Preflight_AllowsCleanReportWorkbookWorkflows() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Clean.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "Beta");
                sheet.CellValue(3, 2, 20d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "Scores", style: TableStyle.TableStyleMedium4);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                ExcelSheet sheet = document.AddWorksheet("Metadata");
                sheet.CellValue(1, 1, "Resource");
                document.Save();
            }

            AddCustomXmlPart(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                Assert.False(report.Can(ExcelPreflightCapability.EditWorkbookStructure));
                Assert.False(report.Can(ExcelPreflightCapability.BindTemplate));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                Assert.Contains("Custom XML parts", string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.EditWorkbookStructure)));
                Assert.Contains("Custom XML parts", string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport)));

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                    report.EnsureCan(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("ExportPdfReport", exception.Message);
                Assert.Contains("Custom XML parts", exception.Message);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsPdfExportForExternalHyperlinks() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.ExternalHyperlink.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.SetHyperlink(2, 1, "https://example.org/spec", display: "Spec");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External hyperlinks");
                Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "External workbook links");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForRelativeExternalHyperlinks() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.RelativeHyperlink.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.CellValue(2, 1, "Spec");
                document.Save();
            }

            AddWorksheetHyperlink(filePath, "A2", "../docs/spec.pdf", UriKind.Relative);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "PDF-unsupported hyperlinks");

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported hyperlinks", diagnostics);
                Assert.Contains("../docs/spec.pdf", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForInternalHyperlinksToSkippedTargets() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.InternalHiddenHyperlink.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Resource");
                summary.CellValue(2, 1, "Hidden target");

                ExcelSheet hidden = document.AddWorksheet("Hidden Details");
                hidden.CellValue(1, 1, "Target");
                hidden.SetHidden(true);
                document.Save();
            }

            AddWorksheetInternalHyperlink(filePath, "A2", "'Hidden Details'!A1");

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "PDF-unsupported hyperlinks");

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported hyperlinks", diagnostics);
                Assert.Contains("Hidden Details", diagnostics);
                Assert.Contains("hidden and skipped", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksFormulaCalculationAndCachedReadsWhenFormulaStateIsUnsafe() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Formulas.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(2, 1, "B");
                sheet.CellFormula(1, 2, "UNIQUE(A1:A2)");
                sheet.CellFormula(2, 2, "B1+1");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                Assert.DoesNotContain("Formula dependency issues", cachedValueDiagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsCachedReadsWhenUnsupportedDependenciesHaveCleanCaches() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.CleanCachedDependency.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, "A");
                sheet.CellFormula(1, 2, "UNIQUE(A1:A1)");
                sheet.CellFormula(1, 3, "B1+1");
                document.Save();
            }

            AddCachedFormulaValue(filePath, "B1", "1");
            AddCachedFormulaValue(filePath, "C1", "2");
            ClearWorkbookRecalculationRequest(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Formula dependency issues");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksCachedReadsWhenWorkbookRequestsRecalculation() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.RecalcRequest.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                Assert.Equal(1, document.Calculate());
                document.ConfigureFullCalculationOnOpen();
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Workbook recalculation requests");

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.Contains("Workbook recalculation requests", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_ReportsRepairHintsForBlockedCapabilities() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.RepairHints.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                document.ConfigureFullCalculationOnOpen();
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                var hints = report.GetRepairHints(ExcelPreflightCapability.UseCachedFormulaValues);
                Assert.Contains(hints, hint => hint.FeatureName == "Missing formula caches");
                Assert.Contains(hints, hint => hint.FeatureName == "Workbook recalculation requests");
                Assert.Contains(hints, hint => hint.Action.Contains("Refresh cached formula values", StringComparison.Ordinal));

                string markdown = report.ToMarkdown();
                Assert.Contains("## Repair Hints", markdown);
                Assert.Contains("Refresh cached formula values", markdown);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsPdfExportWhenUnsafeFormulaCachesAreOnlyHidden() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.HiddenFormulaCaches.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet visible = document.AddWorksheet("Report");
                visible.CellValue(1, 1, "Ready");

                ExcelSheet hidden = document.AddWorksheet("Scratch");
                hidden.CellValue(1, 1, 2d);
                hidden.CellFormula(1, 2, "A1+1");
                hidden.SetHidden(true);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Missing formula caches");
                Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "PDF-missing formula caches");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_IgnoresWorkbookRecalculationRequestWhenNoFormulasExist() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.RecalcNoFormulas.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Ready");
                document.ConfigureFullCalculationOnOpen();
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Workbook recalculation requests");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.UseCachedFormulaValues));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksCellEditsForSignedWorkbooks() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.Signed.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Signed");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Ready");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                Assert.Equal(1, document.Calculate());
                document.InvalidateFormulas();
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                sheet.CellFormula(1, 3, "B1+1");
                sheet.CellFormula(1, 4, "C1+1");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Formula dependency issues");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.CalculateFormulas));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksCircularFormulasWithCleanCachedResults() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.CircularFormulaCaches.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellFormula(1, 1, "B1+1");
                sheet.CellFormula(1, 2, "A1+1");
                document.Save();
            }

            AddCachedFormulaValue(filePath, "A1", "1");
            AddCachedFormulaValue(filePath, "B1", "2");
            ClearWorkbookRecalculationRequest(filePath);

            using ExcelDocument loaded = ExcelDocument.Load(
                filePath,
                new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            ExcelFeatureReport report = loaded.InspectFeatures();

            Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
            Assert.False(report.Can(ExcelPreflightCapability.CalculateFormulas));
            string diagnostics = string.Join(
                Environment.NewLine,
                report.GetCapabilityDiagnostics(ExcelPreflightCapability.CalculateFormulas));
            Assert.Contains("Formula dependency issues", diagnostics);
            Assert.Contains("Circular reference", diagnostics);
            Assert.Contains("Formula dependency issues", Assert.Throws<InvalidOperationException>(() =>
                report.EnsureCan(ExcelPreflightCapability.CalculateFormulas)).Message);
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksUnsupportedFormulasEvenWhenDependenciesOnlyMissCaches() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnsupportedFormulaChain.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                sheet.CellFormula(1, 3, "UNIQUE(B1:B1)");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.CalculateFormulas));
                Assert.Contains("Formula calculation blockers", diagnostics);
                Assert.Contains("UNIQUE", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsPdfExportForCleanCachedUnsupportedFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.CleanCachedUnsupportedFormula.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 7d);
                sheet.CellFormula(1, 2, "UNIQUE(A1:A1)");
                document.Save();
            }

            AddCachedFormulaValue(filePath, "B1", "7");
            ClearWorkbookRecalculationRequest(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.False(report.Can(ExcelPreflightCapability.CalculateFormulas));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Unsupported formulas");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedChartTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.SurfaceChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Charts");
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
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForMixedChartTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.ComboChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Charts");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                    });
                sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220,
                    type: ExcelChartType.ColumnClustered, title: "Combo Chart");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported charts", diagnostics);
                Assert.Contains("mixed per-series chart types", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForSameFamilyMixedChartTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.SameFamilyComboChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Charts");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, ExcelChartType.ColumnClustered),
                        new ExcelChartSeries("Target", new[] { 12d, 18d, 28d }, ExcelChartType.ColumnStacked)
                    });
                sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220,
                    type: ExcelChartType.ColumnClustered, title: "Column Combo Chart");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported charts", diagnostics);
                Assert.Contains("mixed per-series chart types", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnreadableChartSnapshots() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnreadableChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Charts");
                sheet.CellValue(1, 1, "Zone");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 12d);
                sheet.AddChartFromRange("A1:B3", row: 1, column: 5, widthPixels: 320, heightPixels: 180,
                    type: ExcelChartType.ColumnClustered, title: "Literal Chart");
                document.Save();
            }

            RemoveChartRangeFormulas(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unreadable charts", diagnostics);
                Assert.Contains("Literal Chart", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_ReportsDanglingChartPartsWithoutThrowing() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.DanglingChart.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Charts");
                sheet.CellValue(1, 1, "Zone");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 12d);
                ExcelChart chart = sheet.AddChartFromRange("A1:B3", row: 1, column: 5, widthPixels: 320, heightPixels: 180,
                    type: ExcelChartType.ColumnClustered, title: "Dangling Chart");
                chart.Name = "Dangling chart frame";
                document.Save();
            }

            RemoveFirstChartPart(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unreadable charts", diagnostics);
                Assert.Contains("Dangling chart frame", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedWorksheetImages() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnsupportedImage.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.CellValue(1, 1, "Image");
                sheet.AddImage(2, 1, new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, "image/gif",
                    widthPixels: 12, heightPixels: 12, name: "GifLogo");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported images", diagnostics);
                Assert.Contains("image/gif", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedHeaderFooterImages() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnsupportedHeaderImage.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Images");
                sheet.CellValue(1, 1, "Image");
                sheet.SetHeaderImage(HeaderFooterPosition.Center, new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, "image/gif",
                    widthPoints: 24, heightPoints: 24);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported images", diagnostics);
                Assert.Contains("header center", diagnostics);
                Assert.Contains("image/gif", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedHeaderFooterFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnsupportedHeaderFormatting.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Headers");
                sheet.CellValue(1, 1, "Report");
                sheet.SetHeaderFooter(headerCenter: "&UConfidential");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported header/footer formatting", diagnostics);
                Assert.Contains("underline formatting", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForMultiAreaPrintAreas() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.MultiAreaPrintArea.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Top");
                sheet.CellValue(2, 2, "Area one");
                sheet.CellValue(2, 4, "Area two");
                sheet.CellValue(5, 5, "Bottom");
                document.Save();
            }

            AddMultiAreaPrintArea(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported print areas", diagnostics);
                Assert.Contains("multiple print areas", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForPivotTablesAndSparklines() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.PivotSparkline.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 12d);
                sheet.AddPivotTable(
                    sourceRange: "A1:B3",
                    destinationCell: "D1",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", X.DataConsolidateFunctionValues.Sum, "Total Sales") },
                    pivotStyleName: "PivotStyleMedium9");
                sheet.AddSparklines("B2:B3", "C2:C3");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unrendered pivot tables", diagnostics);
                Assert.Contains("PDF-unrendered sparklines", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_IgnoresHiddenSheetPdfOnlyBlockers() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.HiddenPivotSparkline.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet visible = document.AddWorksheet("Report");
                visible.CellValue(1, 1, "Ready");

                ExcelSheet hidden = document.AddWorksheet("Analysis");
                hidden.CellValue(1, 1, "Region");
                hidden.CellValue(1, 2, "Sales");
                hidden.CellValue(2, 1, "North");
                hidden.CellValue(2, 2, 10d);
                hidden.CellValue(3, 1, "South");
                hidden.CellValue(3, 2, 12d);
                hidden.AddPivotTable(
                    sourceRange: "A1:B3",
                    destinationCell: "D1",
                    name: "HiddenPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", X.DataConsolidateFunctionValues.Sum, "Total Sales") });
                hidden.AddSparklines("B2:B3", "C2:C3");
                hidden.SetHidden(true);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Pivot tables");
                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Sparklines");
                Assert.DoesNotContain(report.PartiallyEditableFeatures, feature => feature.Name == "PDF-unrendered pivot tables");
                Assert.DoesNotContain(report.PartiallyEditableFeatures, feature => feature.Name == "PDF-unrendered sparklines");
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForDrawingShapesAndDrawingHyperlinks() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.DrawingShapeLink.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Shapes");
                sheet.CellValue(1, 1, "Callout");
                document.Save();
            }

            AddDrawingShapeAndHyperlink(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unrendered drawing shapes", diagnostics);
                Assert.Contains("Report callout", diagnostics);
                Assert.Contains("Report connector", diagnostics);
                Assert.Contains("PDF-unsupported hyperlinks", diagnostics);
                Assert.Contains("https://example.org/callout", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForRenderableDrawingShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.RenderableDrawingShape.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Shapes");
                sheet.CellValue(1, 1, "Box");
                document.Save();
            }

            AddRenderableDrawingShape(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unrendered drawing shapes", diagnostics);
                Assert.Contains("Renderable rectangle", diagnostics);
                Assert.Contains("shape", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForNonWorksheetSheets() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.ChartSheet.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 10d);
                document.Save();
            }

            AddChartSheet(filePath);

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Non-worksheet sheets");

                string readDiagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ReadWorkbookData));
                Assert.Contains("Non-worksheet sheets", readDiagnostics);

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("Non-worksheet sheets", diagnostics);
                Assert.Contains("ChartOnly", diagnostics);
            }
        }

        private static void AddCustomXmlPart(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            CustomXmlPart part = spreadsheet.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<metadata><owner>OfficeIMO</owner></metadata>"));
            part.FeedData(stream);
        }

        private static void AddCachedFormulaValue(string filePath, string cellReference, string value) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            X.Cell cell = worksheetPart.Worksheet!.Descendants<X.Cell>()
                .First(item => string.Equals(item.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
            cell.CellValue = new X.CellValue(value);
            cell.DataType = X.CellValues.Number;
            cell.CellFormula!.CalculateCell = false;
            worksheetPart.Worksheet.Save();
        }

        private static void AddWorksheetHyperlink(string filePath, string cellReference, string target, UriKind uriKind) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            HyperlinkRelationship relationship = worksheetPart.AddHyperlinkRelationship(new Uri(target, uriKind), true);
            X.Hyperlinks hyperlinks = worksheetPart.Worksheet!.Elements<X.Hyperlinks>().FirstOrDefault()
                ?? worksheetPart.Worksheet.AppendChild(new X.Hyperlinks());
            hyperlinks.Append(new X.Hyperlink {
                Reference = cellReference,
                Id = relationship.Id
            });
            worksheetPart.Worksheet.Save();
        }

        private static void AddWorksheetInternalHyperlink(string filePath, string cellReference, string location) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            X.Hyperlinks hyperlinks = worksheetPart.Worksheet!.Elements<X.Hyperlinks>().FirstOrDefault()
                ?? worksheetPart.Worksheet.AppendChild(new X.Hyperlinks());
            hyperlinks.Append(new X.Hyperlink {
                Reference = cellReference,
                Location = location
            });
            worksheetPart.Worksheet.Save();
        }

        private static void ClearWorkbookRecalculationRequest(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            X.CalculationProperties? properties = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<X.CalculationProperties>();
            if (properties == null) {
                return;
            }

            properties.ForceFullCalculation = false;
            properties.FullCalculationOnLoad = false;
            spreadsheet.WorkbookPart.Workbook.Save();
        }

        private static void AddRenderableDrawingShape(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("1"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("1"),
                        new Xdr.RowOffset("0")),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId("3"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("4"),
                        new Xdr.RowOffset("0")),
                    new Xdr.Shape(
                        new Xdr.NonVisualShapeProperties(
                            new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Renderable rectangle" },
                            new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                        new Xdr.ShapeProperties(
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                        new Xdr.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text("Visible"))))),
                    new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddDrawingShapeAndHyperlink(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.AddHyperlinkRelationship(new Uri("https://example.org/callout", UriKind.Absolute), true, "rIdCalloutLink");

            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("1"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("1"),
                        new Xdr.RowOffset("0")),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId("3"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("4"),
                        new Xdr.RowOffset("0")),
                    new Xdr.Shape(
                        new Xdr.NonVisualShapeProperties(
                            new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Report callout" },
                            new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                        new Xdr.ShapeProperties(
                            new A.PresetGeometry { Preset = A.ShapeTypeValues.RoundRectangle }),
                        new Xdr.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text("Review"))))),
                    new Xdr.ClientData()),
                new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("4"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("1"),
                        new Xdr.RowOffset("0")),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId("5"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("4"),
                        new Xdr.RowOffset("0")),
                    new Xdr.ConnectionShape(
                        new Xdr.NonVisualConnectionShapeProperties(
                            new Xdr.NonVisualDrawingProperties { Id = 3U, Name = "Report connector" },
                            new Xdr.NonVisualConnectorShapeDrawingProperties()),
                        new Xdr.ShapeProperties(
                            new A.PresetGeometry { Preset = A.ShapeTypeValues.Line })),
                    new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddChartSheet(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            ChartsheetPart chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
            chartsheetPart.Chartsheet = new X.Chartsheet();

            X.Sheets sheets = workbookPart.Workbook.Sheets ?? workbookPart.Workbook.AppendChild(new X.Sheets());
            uint nextSheetId = sheets.Elements<X.Sheet>()
                .Select(sheet => sheet.SheetId?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() + 1U;
            sheets.Append(new X.Sheet {
                Id = workbookPart.GetIdOfPart(chartsheetPart),
                SheetId = nextSheetId,
                Name = "ChartOnly"
            });
            chartsheetPart.Chartsheet.Save();
            workbookPart.Workbook.Save();
        }

        private static void AddMultiAreaPrintArea(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            X.Workbook workbook = workbookPart.Workbook;
            workbook.DefinedNames ??= new X.DefinedNames();
            workbook.DefinedNames.Elements<X.DefinedName>()
                .Where(name => string.Equals(name.Name?.Value, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                .ToList()
                .ForEach(name => name.Remove());
            workbook.DefinedNames.Append(new X.DefinedName {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0U,
                Text = "'Report'!$B$2:$B$2,'Report'!$D$2:$D$2"
            });
            workbook.Save();
        }

        private static void RemoveFirstChartPart(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            DrawingsPart drawingsPart = spreadsheet.WorkbookPart!.WorksheetParts
                .Select(part => part.DrawingsPart)
                .First(part => part?.ChartParts.Any() == true)!;
            ChartPart chartPart = drawingsPart.ChartParts.First();
            drawingsPart.DeletePart(chartPart);
            drawingsPart.WorksheetDrawing?.Save();
        }

        private static void RemoveChartRangeFormulas(string filePath) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            ChartPart chartPart = spreadsheet.WorkbookPart!.WorksheetParts
                .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
                .First();
            foreach (C.Formula formula in chartPart.ChartSpace!.Descendants<C.Formula>().ToList()) {
                formula.Remove();
            }

            chartPart.ChartSpace.Save();
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
