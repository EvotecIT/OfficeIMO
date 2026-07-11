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
            var options = new HtmlToWordOptions();
            string html = "<p>Visible before.</p><script>document.body.innerHTML = 'Hidden script';</script><p>Visible after.</p>";

            var document = html.ToWordDocument(options);

            Assert.Equal(new[] { "Visible before.", "Visible after." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Hidden script"));
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("HtmlElementSkipped", diagnostic.Code);
            Assert.Equal("script", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_TemplateContent_IsSkippedWithDiagnostic() {
            var options = new HtmlToWordOptions();
            string html = "<p>Visible.</p><template><p>Hidden template</p></template>";

            var document = html.ToWordDocument(options);

            Assert.Single(document.Paragraphs);
            Assert.Equal("Visible.", document.Paragraphs[0].Text);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("HtmlElementSkipped", diagnostic.Code);
            Assert.Equal("template", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_RawHtmlComments_AreSkippedWithDiagnostic() {
            var callbackDiagnostics = new System.Collections.Generic.List<HtmlConversionDiagnostic>();
            var options = new HtmlToWordOptions {
                DiagnosticHandler = diagnostic => callbackDiagnostics.Add(diagnostic)
            };
            string html = "<p>Visible before.</p><!-- reviewer note --><p>Visible after.</p>";

            var document = html.ToWordDocument(options);

            Assert.Equal(new[] { "Visible before.", "Visible after." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("reviewer note"));
            var diagnostic = Assert.Single(options.Diagnostics, diagnostic => diagnostic.Code == "HtmlCommentSkipped");
            Assert.Equal("comment", diagnostic.Source);
            Assert.Null(diagnostic.Detail);
            Assert.Contains(callbackDiagnostics, diagnostic => diagnostic.Code == "HtmlCommentSkipped");
        }

        [Fact]
        public void HtmlToWord_EmbeddedMediaWidgets_AreSkippedWithDiagnostics() {
            var options = new HtmlToWordOptions();
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

            var document = html.ToWordDocument(options);

            Assert.Equal(new[] { "Before.", "After." }, document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Hidden", System.StringComparison.Ordinal));
            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "HtmlEmbeddedContentSkipped").ToList();
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
            var options = new HtmlToWordOptions();
            string html = "<p>Visible.</p><iframe src=\"https://example.com/widget\">Hidden iframe fallback</iframe><video src=\"movie.mp4\">Hidden video fallback</video>";

            using var document = html.ToWordDocument(options);

            using MemoryStream stream = document.ToDocxStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }
    }
}
