using OfficeIMO.Word.Html;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ScriptContent_IsSkippedWithDiagnostic() {
            var options = new HtmlToWordOptions();
            string html = "<p>Visible before.</p><script>document.body.innerHTML = 'Hidden script';</script><p>Visible after.</p>";

            var document = html.LoadFromHtml(options);

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

            var document = html.LoadFromHtml(options);

            Assert.Single(document.Paragraphs);
            Assert.Equal("Visible.", document.Paragraphs[0].Text);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("HtmlElementSkipped", diagnostic.Code);
            Assert.Equal("template", diagnostic.Source);
        }
    }
}
