using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

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

            VisioDesktopValidationResult result = VisioDesktopValidator.Validate(filePath);

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

            VisioDesktopValidationResult result = VisioDesktopValidator.Validate(filePath, options);

            if (!result.IsAvailable) {
                Assert.False(result.IsValid);
                Assert.NotEmpty(result.Issues);
                return;
            }

            Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues));
            Assert.Contains(roundTripPath, result.OutputFiles);
            string svgPath = Path.Combine(directory, "proof-page1.svg");
            Assert.Contains(svgPath, result.OutputFiles);
            Assert.True(new FileInfo(roundTripPath).Length > 0);
            Assert.True(new FileInfo(svgPath).Length > 0);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
        }

        [Fact]
        public void DesktopValidatorRejectsMissingPathBeforeAutomation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            Assert.Throws<FileNotFoundException>(() => VisioDesktopValidator.Validate(filePath));
        }

        private static bool IsDesktopValidationRequested() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VISIO_DESKTOP_VALIDATION"), "1", StringComparison.Ordinal) ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VISIO_PREMIUM_DESKTOP_BASELINES"), "1", StringComparison.Ordinal) ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);
    }
}
