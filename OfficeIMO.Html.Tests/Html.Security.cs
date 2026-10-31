using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ScriptContent_IsSkippedWithDiagnostic() {
            var options = new WordHtmlImportOptions();
            string html = "<p>Visible before.</p><script>document.body.innerHTML = 'Hidden script';</script><p>Visible after.</p>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
            var document = conversion.Value;

            Assert.Equal(new[] { "Visible before.", "Visible after." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Hidden script"));
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("html-element-skipped", diagnostic.Code);
            Assert.Equal("script", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_TemplateContent_IsSkippedWithDiagnostic() {
            var options = new WordHtmlImportOptions();
            string html = "<p>Visible.</p><template><p>Hidden template</p></template>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
            var document = conversion.Value;

            Assert.Single(document.Paragraphs);
            Assert.Equal("Visible.", document.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("html-element-skipped", diagnostic.Code);
            Assert.Equal("template", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_RawHtmlComments_AreSkippedWithDiagnostic() {
            var options = new WordHtmlImportOptions();
            string html = "<p>Visible before.</p><!-- reviewer note --><p>Visible after.</p>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
            var document = conversion.Value;

            Assert.Equal(new[] { "Visible before.", "Visible after." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("reviewer note"));
            var diagnostic = Assert.Single(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "html-comment-skipped");
            Assert.Equal("comment", diagnostic.Source);
            Assert.Null(diagnostic.Detail);
        }

        [Fact]
        public void HtmlToWord_EmbeddedMediaWidgets_AreSkippedWithDiagnostics() {
            var options = new WordHtmlImportOptions();
            string html = """
                <p>Before.</p>
                <iframe src="https://example.com/widget">Hidden iframe fallback</iframe>
                <object data="movie.swf">Hidden object fallback</object>
                <embed src="movie.swf">
                <video src="movie.mp4">Hidden video fallback</video>
                <audio src="sound.mp3">Hidden audio fallback</audio>
                <canvas>Hidden canvas fallback</canvas>
                <p>After.</p>
                """;

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
            var document = conversion.Value;

            Assert.Equal(new[] { "Before.", "After." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Hidden", System.StringComparison.Ordinal));
            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "html-embedded-content-skipped").ToList();
            Assert.Equal(6, diagnostics.Count);
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "iframe");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "object");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "embed");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "video");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "audio");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "canvas");
        }

        [Fact]
        public void HtmlToWord_EmbeddedMediaWidgets_SavesAsValidOpenXmlDocument() {
            var options = new WordHtmlImportOptions();
            string html = "<p>Visible.</p><iframe src=\"https://example.com/widget\">Hidden iframe fallback</iframe><video src=\"movie.mp4\">Hidden video fallback</video>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
            using var document = conversion.Value;

            using MemoryStream stream = new System.IO.MemoryStream(document.ToBytes());
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }
    }
}
