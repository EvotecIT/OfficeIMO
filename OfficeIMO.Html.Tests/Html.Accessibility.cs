using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlAccessibility {
        [Fact]
        public void HtmlToWord_DocumentLanguage_UsesHtmlLangAttribute() {
            using var document = "<html lang=\"pl-PL\"><body><p>Tekst</p></body></html>".ToWordDocument();

            Assert.Equal("pl-PL", document.Settings.Language);
            Assert.Contains(document.Paragraphs, paragraph => string.Equals(paragraph.Text, "Tekst", StringComparison.Ordinal));
        }

        [Fact]
        public void HtmlToWord_DocumentLanguage_FallsBackToBodyLangAttribute() {
            using var document = "<body lang=\"fr-FR\"><p>Texte</p></body>".ToWordDocument();

            Assert.Equal("fr-FR", document.Settings.Language);
        }

        [Fact]
        public void WordToHtml_DocumentLanguage_ExportsHtmlLangAttribute() {
            using var document = WordDocument.Create();
            document.Settings.Language = "de-DE";
            document.AddParagraph("Text");

            string html = document.ToHtml();

            Assert.Contains("<html lang=\"de-DE\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Text</p>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_ElementLanguage_MapsToRunLanguage() {
            using var document = """
                <p>English <span lang="fr-FR">Bonjour</span> <span xml:lang="pl-PL">Czesc</span></p>
                """.ToWordDocument();

            var runs = document.Paragraphs.Single(paragraph => paragraph.Text.Contains("Bonjour", StringComparison.Ordinal)).GetRuns().ToList();

            Assert.Equal("fr-FR", runs.Single(run => run.Text == "Bonjour").Language);
            Assert.Equal("pl-PL", runs.Single(run => run.Text == "Czesc").Language);
        }

        [Fact]
        public void HtmlToWord_ElementLanguage_InheritsFromContainer() {
            using var document = "<div lang=\"de-DE\">Hallo</div>".ToWordDocument();

            var run = document.Paragraphs.Single(paragraph => paragraph.Text.Contains("Hallo", StringComparison.Ordinal)).GetRuns().Single(run => run.Text == "Hallo");

            Assert.Equal("de-DE", run.Language);
        }

        [Fact]
        public void HtmlToWord_AccessibilityDiagnostics_AreOptIn() {
            var options = new HtmlToWordOptions();

            using var document = "<h1>Title</h1><h3>Skipped</h3><p><img src=\"missing.png\"></p>".ToWordDocument(options);

            Assert.DoesNotContain(options.Diagnostics, diagnostic => diagnostic.Code.StartsWith("Accessibility", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_AccessibilityDiagnostics_ReportMissingImageAltAndWeakLinks() {
            var options = new HtmlToWordOptions {
                EnableAccessibilityDiagnostics = true,
                ImageProcessing = ImageProcessingMode.EmbedDataUriOnly
            };

            using var document = """
                <p>
                  <img src="missing.png">
                  <img src="decorative.png" alt="">
                  <a href="https://example.com/report">click here</a>
                  <a href="https://example.com/empty"><img src="icon.png" alt=""></a>
                  <a href="https://example.com/named" aria-label="Quarterly report"></a>
                </p>
                """.ToWordDocument(options);

            var accessibilityDiagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code.StartsWith("Accessibility", StringComparison.OrdinalIgnoreCase)).ToList();
            Assert.Contains(accessibilityDiagnostics, diagnostic => diagnostic.Code == "AccessibilityImageMissingAlt" && diagnostic.Source == "missing.png");
            Assert.Contains(accessibilityDiagnostics, diagnostic => diagnostic.Code == "AccessibilityLinkTextWeak" && diagnostic.Source == "https://example.com/report");
            Assert.Contains(accessibilityDiagnostics, diagnostic => diagnostic.Code == "AccessibilityLinkTextMissing" && diagnostic.Source == "https://example.com/empty");
            Assert.DoesNotContain(accessibilityDiagnostics, diagnostic => diagnostic.Source == "decorative.png");
            Assert.DoesNotContain(accessibilityDiagnostics, diagnostic => diagnostic.Source == "https://example.com/named");
        }

        [Fact]
        public void HtmlToWord_AccessibilityDiagnostics_ReportHeadingJumpsAndTablesWithoutHeaders() {
            var options = new HtmlToWordOptions {
                EnableAccessibilityDiagnostics = true
            };

            using var document = """
                <h1>Title</h1>
                <h3>Skipped</h3>
                <table id="data">
                  <tr><td>Name</td><td>Total</td></tr>
                  <tr><td>Ada</td><td>42</td></tr>
                </table>
                <table id="headed">
                  <thead><tr><th>Name</th><th>Total</th></tr></thead>
                  <tbody><tr><td>Ada</td><td>42</td></tr></tbody>
                </table>
                """.ToWordDocument(options);

            var accessibilityDiagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code.StartsWith("Accessibility", StringComparison.OrdinalIgnoreCase)).ToList();
            Assert.Contains(accessibilityDiagnostics, diagnostic => diagnostic.Code == "AccessibilityHeadingLevelSkipped" && diagnostic.Source == "h3");
            Assert.Contains(accessibilityDiagnostics, diagnostic => diagnostic.Code == "AccessibilityTableMissingHeader" && diagnostic.Source == "table#data");
            Assert.DoesNotContain(accessibilityDiagnostics, diagnostic => diagnostic.Source == "table#headed");
        }

        [Fact]
        public void HtmlToWord_AccessibilityFixture_SavesAsValidOpenXmlDocument() {
            const string pixelPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
            var options = new HtmlToWordOptions {
                EnableAccessibilityDiagnostics = true
            };

            using var document = $"""
                <html lang="en-US">
                  <body>
                    <main>
                      <h1>Quarterly Accessibility Report</h1>
                      <h2>Summary</h2>
                      <p>Localized note: <span lang="fr-FR">Bonjour</span></p>
                      <p><a href="https://example.com/reports/q1">Download the Q1 accessibility report</a></p>
                      <figure>
                        <img src="data:image/png;base64,{pixelPng}" alt="Status trend icon" width="16" height="16">
                        <figcaption>Status trend icon</figcaption>
                      </figure>
                      <table id="scorecard">
                        <caption>Accessibility scorecard</caption>
                        <thead>
                          <tr><th scope="col">Check</th><th scope="col">Result</th></tr>
                        </thead>
                        <tbody>
                          <tr><td>Images</td><td>Passed</td></tr>
                          <tr><td>Links</td><td>Passed</td></tr>
                        </tbody>
                      </table>
                      <ul>
                        <li><input type="checkbox" checked> Reviewed by document owner</li>
                      </ul>
                    </main>
                  </body>
                </html>
                """.ToWordDocument(options);

            Assert.DoesNotContain(options.Diagnostics, diagnostic => diagnostic.Code.StartsWith("Accessibility", StringComparison.OrdinalIgnoreCase));

            using MemoryStream stream = document.SaveAsMemoryStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }
    }
}
