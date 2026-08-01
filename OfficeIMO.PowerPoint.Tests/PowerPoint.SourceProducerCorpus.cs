using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
                Dictionary<string, string[]> sourceFingerprints = artifact.Focus
                    .ToDictionary(focus => focus,
                        focus => GetFocusEntries(
                            presentation.OpenXmlDocument, focus),
                        StringComparer.OrdinalIgnoreCase);
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
                foreach (string focus in artifact.Focus) {
                    string[] reopenedFingerprint = GetFocusEntries(
                        reopened.OpenXmlDocument, focus);
                    bool preserved = string.Equals(focus,
                            "extension preservation",
                            StringComparison.OrdinalIgnoreCase)
                        ? ContainsAllEntries(reopenedFingerprint,
                            sourceFingerprints[focus])
                        : sourceFingerprints[focus].SequenceEqual(
                            reopenedFingerprint, StringComparer.Ordinal);
                    Assert.True(preserved,
                        artifact.File + " did not preserve focus '" + focus
                        + "'. Expected " + ComputeSha256(sourceFingerprints[focus])
                        + " but found " + ComputeSha256(reopenedFingerprint) + ".");
                }
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

        private static string[] GetFocusEntries(PresentationDocument document,
            string focus) {
            PresentationPart presentation = document.PresentationPart
                ?? throw new InvalidDataException("Corpus presentation has no presentation part.");
            string normalized = focus.Trim().ToLowerInvariant();
            IEnumerable<string> entries;
            switch (normalized) {
                case "blank template":
                    entries = new[] {
                        "slides=" + presentation.SlideParts.Count(),
                        "size=" + presentation.Presentation?.SlideSize?.OuterXml
                    };
                    break;
                case "master and layout preservation":
                case "theme and layout inheritance":
                    entries = presentation.SlideMasterParts.Select(master =>
                            "master|" + master.Uri + "|" + master.SlideMaster.OuterXml)
                        .Concat(presentation.SlideMasterParts.SelectMany(master =>
                            master.SlideLayoutParts.Select(layout =>
                                "layout|" + layout.Uri + "|" + layout.SlideLayout.OuterXml)))
                        .Concat(GetAllParts(presentation).OfType<ThemePart>()
                            .Select(theme => "theme|" + theme.Uri + "|"
                                + theme.Theme.OuterXml));
                    break;
                case "themes":
                    entries = GetAllParts(presentation).OfType<ThemePart>()
                        .Select(theme => "theme|" + theme.Uri + "|"
                            + theme.Theme.OuterXml);
                    break;
                case "tables":
                case "powerpoint-authored tables":
                case "table style preservation":
                    entries = presentation.SlideParts.SelectMany(slide =>
                        slide.Slide.Descendants<A.Table>().Select(table =>
                            slide.Uri + "|" + table.OuterXml));
                    break;
                case "charts":
                    entries = presentation.SlideParts.SelectMany(slide =>
                            slide.ChartParts.Select(chart =>
                                "chart|" + chart.Uri + "|" + chart.ChartSpace.OuterXml))
                        .Concat(presentation.SlideParts.SelectMany(slide =>
                            slide.ChartParts.SelectMany(chart =>
                                chart.GetPartsOfType<EmbeddedPackagePart>())
                            .Select(package => "workbook|" + package.Uri + "|"
                                + ComputePartSha256(package))));
                    break;
                case "pictures":
                    entries = presentation.SlideParts.SelectMany(slide =>
                            slide.Slide.Descendants<P.Picture>().Select(picture =>
                                "picture|" + slide.Uri + "|" + picture.OuterXml))
                        .Concat(GetAllParts(presentation).OfType<ImagePart>()
                            .Select(image => "image|" + image.Uri + "|"
                                + ComputePartSha256(image)));
                    break;
                case "multi-layout editing":
                    entries = presentation.SlideParts.Select(slide =>
                        "slide|" + slide.Uri + "|layout="
                        + (slide.SlideLayoutPart?.Uri.ToString() ?? string.Empty));
                    break;
                case "title placeholder":
                    entries = presentation.SlideParts.SelectMany(slide =>
                        slide.Slide.Descendants<P.Shape>()
                            .Where(IsTitlePlaceholder)
                            .Select(shape => slide.Uri + "|" + shape.OuterXml));
                    break;
                case "transitions":
                    entries = presentation.SlideParts.Select(slide =>
                        slide.Slide.Transition == null
                            ? slide.Uri + "|none"
                            : slide.Uri + "|" + slide.Slide.Transition.OuterXml);
                    break;
                case "extension preservation":
                    entries = GetExtensionEntries(presentation);
                    break;
                default:
                    throw new InvalidDataException(
                        "Corpus focus has no semantic fingerprint: " + focus);
            }
            return entries.OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
        }

        private static bool ContainsAllEntries(IEnumerable<string> actual,
            IEnumerable<string> expected) {
            Dictionary<string, int> counts = actual.GroupBy(value => value,
                    StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(),
                    StringComparer.Ordinal);
            foreach (string entry in expected) {
                if (!counts.TryGetValue(entry, out int count) || count == 0) {
                    return false;
                }
                counts[entry] = count - 1;
            }
            return true;
        }

        private static bool IsTitlePlaceholder(P.Shape shape) {
            P.PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();
            return placeholder?.Type?.Value == P.PlaceholderValues.Title
                || placeholder?.Type?.Value == P.PlaceholderValues.CenteredTitle;
        }

        private static IEnumerable<string> GetExtensionEntries(
            PresentationPart presentation) {
            foreach (OpenXmlPart part in GetAllParts(presentation)) {
                if (part.ContentType.IndexOf("xml",
                        StringComparison.OrdinalIgnoreCase) < 0) continue;
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                XDocument document;
                try {
                    document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                } catch (XmlException) {
                    continue;
                }
                foreach (XElement extension in document.Descendants()
                             .Where(element => element.Name.LocalName == "ext")) {
                    yield return part.Uri + "|" + extension.ToString(
                        SaveOptions.DisableFormatting);
                }
            }
        }

        private static IReadOnlyList<OpenXmlPart> GetAllParts(
            PresentationPart presentation) {
            var result = new List<OpenXmlPart>();
            var pending = new Queue<OpenXmlPart>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pending.Enqueue(presentation);
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Dequeue();
                if (!visited.Add(part.Uri.ToString())) continue;
                result.Add(part);
                foreach (IdPartPair child in part.Parts) {
                    pending.Enqueue(child.OpenXmlPart);
                }
            }
            return result;
        }

        private static string ComputePartSha256(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using SHA256 sha256 = SHA256.Create();
            return ToHex(sha256.ComputeHash(stream));
        }

        private static string ComputeSha256(IEnumerable<string> entries) {
            using SHA256 sha256 = SHA256.Create();
            byte[] bytes = Encoding.UTF8.GetBytes(string.Join("\n", entries));
            return ToHex(sha256.ComputeHash(bytes));
        }

        private static string ToHex(IEnumerable<byte> bytes) =>
            string.Concat(bytes.Select(value => value.ToString("x2")));

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
