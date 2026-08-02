using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests {
    public class VisioDesktopValidation {
        [Fact]
        public void DesktopValidatorReportsAvailabilityOrOpensGeneratedDocument() {
            if (!IsDesktopValidationRequested()) {
                return;
            }

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));
            document.Save();

            VisioDesktopValidationResult result = VisioDesktopBaselineValidator.Validate(filePath);

            if (!result.IsAvailable) {
                Assert.False(result.IsValid);
                Assert.NotEmpty(result.Issues);
                Assert.Contains(result.Issues, issue => issue.Contains("not available", StringComparison.OrdinalIgnoreCase));
                return;
            }

            Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues));
            Assert.Empty(result.Issues);
        }

        [Fact]
        public void DesktopValidatorCanRoundTripAndExportGeneratedDocument() {
            if (!IsDesktopValidationRequested()) {
                return;
            }

            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-VisioDesktop-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string filePath = Path.Combine(directory, "source.vsdx");
            string roundTripPath = Path.Combine(directory, "roundtrip.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = page.AddRectangle(2, 2, 2, 1, "Start");
            VisioShape end = page.AddRectangle(5, 2, 2, 1, "End");
            page.AddConnector(start, end, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            document.Save();

            VisioDesktopValidationOptions options = VisioDesktopValidationOptions.RoundTripWithSvg();
            options.SaveCopyPath = roundTripPath;
            options.ExportDirectory = directory;
            options.ExportFileNamePrefix = "proof";
            options.ExportFormats.Add(VisioDesktopExportFormat.Pdf);

            VisioDesktopValidationResult result = VisioDesktopBaselineValidator.Validate(filePath, options);

            if (!result.IsAvailable) {
                Assert.False(result.IsValid);
                Assert.NotEmpty(result.Issues);
                return;
            }

            Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues));
            Assert.Contains(roundTripPath, result.OutputFiles);
            string svgPath = Path.Combine(directory, "proof-page1.svg");
            string pdfPath = Path.Combine(directory, "proof-page1.pdf");
            Assert.Contains(svgPath, result.OutputFiles);
            Assert.Contains(pdfPath, result.OutputFiles);
            Assert.True(new FileInfo(roundTripPath).Length > 0);
            Assert.True(new FileInfo(svgPath).Length > 0);
            Assert.True(new FileInfo(pdfPath).Length > 0);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
        }

        [Fact]
        public void DesktopValidatorRejectsMissingPathBeforeAutomation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            Assert.Throws<FileNotFoundException>(() => VisioDesktopBaselineValidator.Validate(filePath));
        }

        [Fact]
        public void DesktopValidatorRejectsStructurallyInvalidReferenceOutputs() {
            string directory = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-VisioDesktopOutputValidation-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                string png = Path.Combine(directory, "valid.png");
                File.WriteAllBytes(png, VisualBaselineTestSupport.CreateRgbPng(
                    1, 1, new byte[] { 12, 34, 56 }));
                Assert.True(VisioDesktopBaselineValidator.ValidateOutputFile(
                    png, out string validIssue), validIssue);

                string invalidPng = Path.Combine(directory, "invalid.png");
                File.WriteAllText(invalidPng, "not a PNG");
                Assert.False(VisioDesktopBaselineValidator.ValidateOutputFile(
                    invalidPng, out string invalidIssue));
                Assert.Contains("PNG", invalidIssue, StringComparison.Ordinal);

                string invalidSvg = Path.Combine(directory, "invalid.svg");
                File.WriteAllText(invalidSvg, "<html />");
                Assert.False(VisioDesktopBaselineValidator.ValidateOutputFile(
                    invalidSvg, out string svgIssue));
                Assert.Contains("SVG root", svgIssue, StringComparison.Ordinal);

                string emptySvg = Path.Combine(directory, "empty.svg");
                File.WriteAllText(emptySvg,
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\" />");
                Assert.False(VisioDesktopBaselineValidator.ValidateOutputFile(
                    emptySvg, out string emptySvgIssue));
                Assert.Contains("visible graphical content", emptySvgIssue,
                    StringComparison.OrdinalIgnoreCase);

                string validSvg = Path.Combine(directory, "valid.svg");
                File.WriteAllText(validSvg,
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><rect width=\"10\" height=\"10\" /></svg>");
                Assert.True(VisioDesktopBaselineValidator.ValidateOutputFile(
                    validSvg, out string validSvgIssue), validSvgIssue);

                string validPdf = Path.Combine(directory, "valid.pdf");
                byte[] pdf = PdfCore.PdfDocument.Create()
                    .Paragraph(paragraph => paragraph.Text("Desktop proof"))
                    .ToBytes();
                File.WriteAllBytes(validPdf, pdf);
                Assert.True(VisioDesktopBaselineValidator.ValidateOutputFile(
                    validPdf, out string validPdfIssue, expectedPdfPageCount: 1),
                    validPdfIssue);
                Assert.False(VisioDesktopBaselineValidator.ValidateOutputFile(
                    validPdf, out string pageCountIssue, expectedPdfPageCount: 2));
                Assert.Contains("expected 2", pageCountIssue,
                    StringComparison.OrdinalIgnoreCase);

                string invalidPdf = Path.Combine(directory, "invalid.pdf");
                File.WriteAllBytes(invalidPdf, pdf.Take(pdf.Length / 2).ToArray());
                Assert.False(VisioDesktopBaselineValidator.ValidateOutputFile(
                    invalidPdf, out string invalidPdfIssue));
                Assert.NotEmpty(invalidPdfIssue);
            } finally {
                if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            }
        }

        private static bool IsDesktopValidationRequested() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VISIO_DESKTOP_VALIDATION"), "1", StringComparison.Ordinal) ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VISIO_PREMIUM_DESKTOP_BASELINES"), "1", StringComparison.Ordinal) ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);
    }
}
