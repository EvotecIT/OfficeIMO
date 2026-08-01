using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointSourceProducerCorpusTests {
        [Fact]
        public void MicrosoftPowerPointCorpusIsCompleteAndSurvivesOpenEditSave() {
            string root = GetRepositoryRoot();
            string corpusDirectory = Path.Combine(root, "Assets", "PowerPointTemplates");
            string manifestPath = Path.Combine(corpusDirectory, "corpus-manifest.json");
            PowerPointSourceCorpusManifest manifest = JsonSerializer.Deserialize<
                PowerPointSourceCorpusManifest>(File.ReadAllText(manifestPath),
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new InvalidDataException("PowerPoint source corpus manifest is empty.");

            Assert.Equal(1, manifest.SchemaVersion);
            Assert.Equal("Microsoft Office PowerPoint 16", manifest.Producer);
            Assert.NotEmpty(manifest.Artifacts);
            string[] actualFiles = Directory.GetFiles(corpusDirectory, "*.pptx")
                .Select(Path.GetFileName)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray()!;
            string[] declaredFiles = manifest.Artifacts.Select(artifact => artifact.File)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            Assert.Equal(declaredFiles, actualFiles,
                StringComparer.OrdinalIgnoreCase);

            foreach (PowerPointSourceCorpusArtifact artifact in manifest.Artifacts) {
                Assert.NotEmpty(artifact.Focus);
                string sourcePath = Path.Combine(corpusDirectory, artifact.File);
                Assert.True(File.Exists(sourcePath), "Missing corpus fixture: " + sourcePath);
                Assert.Equal(artifact.Sha256, ComputeSha256(sourcePath));

                using PowerPointPresentation presentation =
                    PowerPointPresentation.Load(sourcePath);
                AssertMicrosoftPowerPointProducer(presentation.OpenXmlDocument);
                string[] sourceValidation = GetValidationSignature(
                    presentation.ValidateDocument());
                PowerPointSlide editedSlide = presentation.Slides.Count == 0
                    ? presentation.AddSlide()
                    : presentation.Slides[0];
                editedSlide.AddTextBoxPoints("OfficeIMO corpus edit", 18, 18, 180, 24);
                using var editedBytes = new MemoryStream();
                presentation.Save(editedBytes);

                editedBytes.Position = 0;
                using PowerPointPresentation reopened =
                    PowerPointPresentation.Load(editedBytes);
                Assert.Equal(presentation.Slides.Count, reopened.Slides.Count);
                Assert.Contains(reopened.Slides[0].TextBoxes,
                    textBox => textBox.Text == "OfficeIMO corpus edit");
                Assert.Equal(sourceValidation,
                    GetValidationSignature(reopened.ValidateDocument()));
                OfficeImageExportResult firstSlide = reopened.Slides[0].ExportImage(
                    OfficeImageExportFormat.Png);
                Assert.DoesNotContain(firstSlide.Diagnostics,
                    diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
                Assert.True(OfficePngReader.TryDecode(firstSlide.Bytes,
                    out OfficeRasterImage? raster));
                Assert.True(raster!.Width > 0 && raster.Height > 0);
            }
        }

        private static void AssertMicrosoftPowerPointProducer(
            PresentationDocument document) {
            string? application = document.ExtendedFilePropertiesPart?
                .Properties?.Application?.Text;
            string? version = document.ExtendedFilePropertiesPart?
                .Properties?.ApplicationVersion?.Text;
            Assert.Equal("Microsoft Office PowerPoint", application);
            Assert.StartsWith("16.", version, StringComparison.Ordinal);
        }

        private static string ComputeSha256(string path) {
            using SHA256 sha256 = SHA256.Create();
            using FileStream stream = File.OpenRead(path);
            byte[] hash = sha256.ComputeHash(stream);
            return string.Concat(hash.Select(value => value.ToString("x2")));
        }

        private static string[] GetValidationSignature(
            IEnumerable<ValidationErrorInfo> errors) {
            return errors.Select(error => string.Join("|",
                    error.Id ?? string.Empty,
                    error.ErrorType.ToString(),
                    error.Part?.Uri.ToString() ?? string.Empty,
                    error.Path?.XPath ?? string.Empty,
                    error.Description ?? string.Empty))
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
        }

        private static string GetRepositoryRoot() {
            DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                    return directory.FullName;
                }
                directory = directory.Parent;
            }
            throw new DirectoryNotFoundException(
                "Could not locate the OfficeIMO repository root.");
        }
    }

    internal sealed class PowerPointSourceCorpusManifest {
        public int SchemaVersion { get; set; }
        public string Producer { get; set; } = string.Empty;
        public List<PowerPointSourceCorpusArtifact> Artifacts { get; set; } = new();
    }

    internal sealed class PowerPointSourceCorpusArtifact {
        public string File { get; set; } = string.Empty;
        public string Sha256 { get; set; } = string.Empty;
        public List<string> Focus { get; set; } = new();
    }
}
