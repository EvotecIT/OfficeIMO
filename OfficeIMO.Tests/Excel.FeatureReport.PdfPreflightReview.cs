using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForPngBytesRejectedByPdfWriter() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.InvalidPngImage.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Images");
                sheet.CellValue(1, 1, "Image");
                sheet.AddImage(2, 1, CreatePngWithInvalidCrc(), "image/png", widthPixels: 12, heightPixels: 12, name: "InvalidPng");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported images", diagnostics);
                Assert.Contains("invalid CRC", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForMixedHeaderFooterStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.MixedHeaderFormatting.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Headers");
                sheet.CellValue(1, 1, "Report");
                sheet.SetHeaderFooter(headerLeft: "&BTotal", headerCenter: "Page &P");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported header/footer formatting", diagnostics);
                Assert.Contains("mixed header/footer formatting", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForUnsupportedHeaderFooterFonts() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.UnsupportedHeaderFont.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Headers");
                sheet.CellValue(1, 1, "Report");
                sheet.SetHeaderFooter(headerCenter: "&\"UnmappedFont\"Report");
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported header/footer formatting", diagnostics);
                Assert.Contains("unsupported font formatting", diagnostics);
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsPdfExportWhenUnsafeFormulaCachesAreOutsidePrintArea() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.FormulaOutsidePrintArea.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);
                sheet.CellFormula(5, 4, "B2+1");
                document.SetPrintArea(sheet, "A1:B2", save: false);
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Missing formula caches");
                Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "PDF-missing formula caches");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_AllowsPdfExportWhenUnsafeFormulaCachesAreHiddenRows() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.HiddenRowFormulaCache.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Ready");
                sheet.CellFormula(5, 1, "A1+1");
                sheet.SetRowHidden(5, true);
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.UseCachedFormulaValues));
                Assert.True(report.Can(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Missing formula caches");
                Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "PDF-missing formula caches");
                Assert.Empty(report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
            }
        }

        [Fact]
        public void FeatureReport_Preflight_BlocksPdfExportForPrintTitleColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "FeatureReport.Preflight.PrintTitleColumns.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 1, lastCol: 1, save: false);
                document.Save(false);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                string diagnostics = string.Join(Environment.NewLine,
                    report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                Assert.Contains("PDF-unsupported print titles", diagnostics);
                Assert.Contains("print-title columns", diagnostics);
            }
        }

        private static byte[] CreatePngWithInvalidCrc() {
            return new byte[] {
                137, 80, 78, 71, 13, 10, 26, 10,
                0, 0, 0, 13,
                73, 72, 68, 82,
                0, 0, 0, 1,
                0, 0, 0, 1,
                8, 2, 0, 0, 0,
                0, 0, 0, 0
            };
        }
    }
}
