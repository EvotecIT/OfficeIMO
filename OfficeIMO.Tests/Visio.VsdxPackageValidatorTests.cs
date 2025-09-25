using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioVsdxPackageValidatorTests {
        private static string CreateSampleVisioDocument() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(path);
            VisioPage page = document.AddPage("Page-1", 8.5, 11, VisioMeasurementUnit.Inches);
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Sample"));
            document.Save();

            return path;
        }

        [Fact]
        public void ValidateFile_FlagsMissingPagesOverride_In2012Package() {
            string tempPath = CreateSampleVisioDocument();
            try {
                var validator = new VsdxPackageValidator();
                bool result = validator.ValidateFile(tempPath);

                Assert.False(result);
                Assert.Contains("Missing override for /visio/pages/pages.xml", validator.Errors);
            } finally {
                if (File.Exists(tempPath)) {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void ValidateFileStreaming_Passes_On2012Package() {
            string tempPath = CreateSampleVisioDocument();
            try {
                var validator = new VsdxPackageValidator();
                bool result = validator.ValidateFileStreaming(tempPath);

                Assert.True(result, string.Join(Environment.NewLine, validator.Errors));
                Assert.Empty(validator.Errors);
            } finally {
                if (File.Exists(tempPath)) {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void FixFile_CreatesCorrectedPackage_For2012Document() {
            string inputPath = CreateSampleVisioDocument();
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            try {
                var validator = new VsdxPackageValidator();
                bool fixedResult = validator.FixFile(inputPath, outputPath);

                Assert.True(fixedResult);
                Assert.True(File.Exists(outputPath));

                var postValidator = new VsdxPackageValidator();
                bool postResult = postValidator.ValidateFile(outputPath);
                Assert.True(postResult, string.Join(Environment.NewLine, postValidator.Errors));
                Assert.Empty(postValidator.Errors);
            } finally {
                if (File.Exists(inputPath)) {
                    File.Delete(inputPath);
                }
                if (File.Exists(outputPath)) {
                    File.Delete(outputPath);
                }
            }
        }
    }
}

