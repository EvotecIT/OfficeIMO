using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Html;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Capability-gallery helpers for HTML-to-Word-to-HTML roundtrip proof artifacts.
    /// </summary>
    public static class WordHtmlCapabilityGalleryExtensions {
        private const string WordDocumentMediaType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        /// <summary>
        /// Saves source HTML, generated DOCX, roundtrip HTML, and shared manifest files for a Word HTML roundtrip scenario.
        /// </summary>
        /// <param name="source">Parsed source HTML to import into Word.</param>
        /// <param name="directoryPath">Directory where artifacts will be written.</param>
        /// <param name="options">Optional gallery options.</param>
        /// <returns>Shared gallery manifest describing the generated artifacts and roundtrip expectations.</returns>
        public static HtmlCapabilityGalleryManifest SaveHtmlCapabilityGallery(this HtmlConversionDocument source, string directoryPath, WordHtmlCapabilityGalleryOptions? options = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Artifact directory cannot be empty.", nameof(directoryPath));

            options ??= new WordHtmlCapabilityGalleryOptions();
            Directory.CreateDirectory(directoryPath);

            string filePrefix = ToFilePrefix(options.ScenarioId);
            string inputPath = Path.Combine(directoryPath, filePrefix + ".input.html");
            string docxPath = Path.Combine(directoryPath, filePrefix + ".docx");
            string roundTripPath = Path.Combine(directoryPath, filePrefix + ".roundtrip.html");
            string manifestPath = Path.Combine(directoryPath, filePrefix + ".manifest.md");
            string manifestJsonPath = Path.Combine(directoryPath, filePrefix + ".manifest.json");

            HtmlToWordOptions importOptions = CreateImportOptions(options.ImportOptions);
            WordToHtmlOptions exportOptions = CreateExportOptions(options.ExportOptions);
            HtmlToWordResult importResult = source.ToWordDocumentResult(importOptions);
            using WordDocument document = importResult.RequireValue();
            using MemoryStream packageStream = document.ToStream();
            string roundTripHtml = document.ToHtml(exportOptions);

            HtmlCapabilityGalleryArtifact sourceArtifact = HtmlCapabilityGalleryArtifact.WriteTextFile("source", "input-html", inputPath, "text/html", source.SourceHtml);
            OfficeFileCommit.WriteAllBytes(docxPath, packageStream.ToArray());
            HtmlCapabilityGalleryArtifact docxArtifact = HtmlCapabilityGalleryArtifact.FromFile("docx", "docx", docxPath, WordDocumentMediaType);
            HtmlCapabilityGalleryArtifact roundTripArtifact = HtmlCapabilityGalleryArtifact.WriteTextFile("roundtrip", "roundtrip-html", roundTripPath, "text/html", roundTripHtml);

            var scenario = new HtmlCapabilityGalleryScenario(
                options.ScenarioId,
                options.Title,
                "Word HTML",
                "Validates HTML import, DOCX package validity, round-trip HTML export, form controls, tables, resources, and diagnostics.");
            var result = new HtmlCapabilityGalleryResult(scenario);
            result.AddArtifact(sourceArtifact);
            result.AddArtifact(docxArtifact);
            result.AddArtifact(roundTripArtifact);
            result.Diagnostics.AddRange(importResult.Report.Diagnostics);
            AppendOpenXmlValidationDiagnostics(result, packageStream);

            HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(source.SourceHtml, roundTripHtml);
            HtmlResourceManifest resourceManifest = HtmlResourcePipeline.BuildManifest(source.SourceHtml, new HtmlResourcePipelineOptions {
                ResourceUrlPolicy = options.ResourceUrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()
            });
            var manifest = new HtmlCapabilityGalleryManifest(
                result,
                source.ProfileContract.Profile,
                score,
                resourceManifest,
                CreateExpectations(source.SourceHtml),
                new[] { OfficeHtmlConversionProfile.WordDocumentRoundTrip });

            string manifestMarkdown = HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest);
            string manifestJson = HtmlCapabilityGalleryManifestJsonWriter.ToJson(manifest);
            HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-md", "manifest-markdown", manifestPath, "text/markdown", manifestMarkdown);
            HtmlCapabilityGalleryArtifact.WriteTextFile("manifest-json", "manifest-json", manifestJsonPath, "application/json", manifestJson);

            return manifest;
        }

        private static HtmlToWordOptions CreateImportOptions(HtmlToWordOptions? options) {
            HtmlToWordOptions importOptions = options?.Clone() ?? HtmlToWordOptions.CreateTrustedDocumentProfile();
            importOptions.EnableAccessibilityDiagnostics = true;
            return importOptions;
        }

        private static WordToHtmlOptions CreateExportOptions(WordToHtmlOptions? options) {
            if (options != null) {
                return options;
            }

            return new WordToHtmlOptions {
                IncludeListStyles = true,
                IncludeTableColumnGroups = true,
                IncludeDefaultCss = true,
                ExportFootnotes = true,
                ExportEndnotes = true
            };
        }

        private static void AppendOpenXmlValidationDiagnostics(HtmlCapabilityGalleryResult result, MemoryStream packageStream) {
            packageStream.Position = 0;
            using WordprocessingDocument package = WordprocessingDocument.Open(packageStream, false);
            IReadOnlyList<ValidationErrorInfo> errors = new OpenXmlValidator().Validate(package).ToList();
            if (errors.Count == 0) {
                result.Diagnostics.Add(
                    "OfficeIMO.Word.Html",
                    "WordOpenXmlPackageValid",
                    "Generated DOCX package passed OpenXML validation.",
                    HtmlDiagnosticSeverity.Info);
                return;
            }

            foreach (ValidationErrorInfo error in errors) {
                result.Diagnostics.Add(
                    "OfficeIMO.Word.Html",
                    "WordOpenXmlValidationError",
                    error.Description ?? "Generated DOCX package failed OpenXML validation.",
                    HtmlDiagnosticSeverity.Error,
                    error.Path?.XPath,
                    error.Id);
            }
        }

        private static IReadOnlyList<HtmlCapabilityGalleryExpectation> CreateExpectations(string html) {
            var expectations = new List<HtmlCapabilityGalleryExpectation>();
            if (ContainsAny(html, "<h1", "<h2", "<h3", "<h4", "<h5", "<h6")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("headings", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains heading elements"));
            }

            if (ContainsAny(html, "<table")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("tables", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains table content"));
            }

            if (ContainsAny(html, "<thead", "<tbody", "<tfoot")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("table sections", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains table section elements"));
            }

            if (ContainsAny(html, "<form", "<input", "<select", "<textarea", "<button")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("form controls", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains form control elements"));
            }

            if (ContainsAny(html, "<img", "<picture", "<svg")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("images", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains image or SVG evidence"));
            }

            if (ContainsAny(html, "<!--")) {
                expectations.Add(new HtmlCapabilityGalleryExpectation("comments", HtmlCapabilityGalleryExpectationOutcome.Reported, "HtmlCommentSkipped diagnostic is present"));
            }

            expectations.Add(new HtmlCapabilityGalleryExpectation("docx package", HtmlCapabilityGalleryExpectationOutcome.Preserved, "generated DOCX passes OpenXML validation"));
            return expectations;
        }

        private static bool ContainsAny(string value, params string[] needles) {
            foreach (string needle in needles) {
                if (value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) {
                    return true;
                }
            }

            return false;
        }

        private static string ToFilePrefix(string scenarioId) {
            var builder = new StringBuilder();
            foreach (char ch in scenarioId ?? string.Empty) {
                builder.Append(char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '-');
            }

            return builder.Length == 0 ? "word-html-gallery" : builder.ToString();
        }
    }
}
