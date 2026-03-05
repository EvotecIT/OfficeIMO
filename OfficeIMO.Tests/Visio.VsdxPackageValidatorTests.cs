using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
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

        [Fact]
        public void FixFileStreaming_PreservesNonPageDocumentRelationships() {
            string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            try {
                VisioDocument document = VisioDocument.Create(inputPath);
                document.UseMastersByDefault = true;
                document.Theme = new VisioTheme { Name = "Validator Theme" };
                VisioPage page = document.AddPage("Page-1", 8.5, 11, VisioMeasurementUnit.Inches);
                page.AddRectangle(1, 1, 2, 1, "Sample");
                document.Save();

                var validator = new VsdxPackageValidator();
                bool fixedResult = validator.FixFileStreaming(inputPath, outputPath);

                Assert.True(fixedResult, string.Join(Environment.NewLine, validator.Errors));

                using ZipArchive archive = ZipFile.OpenRead(outputPath);
                ZipArchiveEntry? relsEntry = archive.GetEntry("visio/_rels/document.xml.rels");
                Assert.NotNull(relsEntry);
                using Stream relsStream = relsEntry!.Open();
                XDocument relsXml = XDocument.Load(relsStream);
                XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
                string[] types = relsXml.Root!.Elements(pr + "Relationship")
                    .Select(e => (string?)e.Attribute("Type") ?? string.Empty)
                    .ToArray();

                Assert.Contains("http://schemas.microsoft.com/visio/2010/relationships/pages", types);
                Assert.Contains("http://schemas.microsoft.com/visio/2010/relationships/windows", types);
                Assert.Contains("http://schemas.microsoft.com/visio/2010/relationships/theme", types);
                Assert.Contains("http://schemas.microsoft.com/visio/2010/relationships/masters", types);
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

