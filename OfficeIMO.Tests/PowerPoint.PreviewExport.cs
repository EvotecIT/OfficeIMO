using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPreviewExportTests {
        [Fact]
        public void PreviewExporterValidatesInputBeforeAutomation() {
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
            string existingPresentation = System.IO.Path.Combine(outputDirectory, "existing.pptx");

            try {
                Directory.CreateDirectory(outputDirectory);
                File.WriteAllBytes(existingPresentation, Array.Empty<byte>());

                Assert.Throws<FileNotFoundException>(() =>
                    PowerPointPreviewExporter.TryExportSlides(
                        System.IO.Path.Combine(outputDirectory, "missing.pptx"),
                        outputDirectory,
                        out _));
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    PowerPointPreviewExporter.TryExportSlides(
                        existingPresentation,
                        outputDirectory,
                        out _,
                        width: -1));
            } finally {
                if (Directory.Exists(outputDirectory)) {
                    Directory.Delete(outputDirectory, recursive: true);
                }
            }
        }

        [Fact]
        public void PreviewExporterDoesNotOpenExistingPresentationWithoutTrustOptIn() {
            string outputDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
            string existingPresentation = System.IO.Path.Combine(outputDirectory, "existing.pptx");

            try {
                Directory.CreateDirectory(outputDirectory);
                File.WriteAllBytes(existingPresentation, Array.Empty<byte>());

                bool exported = PowerPointPreviewExporter.TryExportSlides(
                    existingPresentation,
                    outputDirectory,
                    out PowerPointPreviewExportResult result);

                Assert.False(exported);
                Assert.False(result.Succeeded);
                Assert.Empty(result.Files);
                Assert.Null(result.Exception);
                Assert.Contains("TrustPresentationFile", result.Message, StringComparison.Ordinal);
            } finally {
                if (Directory.Exists(outputDirectory)) {
                    Directory.Delete(outputDirectory, recursive: true);
                }
            }
        }

        [Fact]
        public void PreviewExporterReportsAvailabilityWithoutThrowing() {
            bool available = PowerPointPreviewExporter.IsPowerPointAutomationAvailable();
            Assert.IsType<bool>(available);
        }
    }
}
