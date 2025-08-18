using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlAdditionalHeadTags {
        [Fact]
        public void Test_WordToHtml_WithAdditionalMetaTags() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Content");

            var options = new WordToHtmlOptions();
            options.AdditionalMetaTags.Add(("viewport", "width=device-width, initial-scale=1"));

            string html = doc.ToHtml(options);

            Assert.Contains("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_WithAdditionalLinkTags() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Content");

            var options = new WordToHtmlOptions();
            options.AdditionalLinkTags.Add(("stylesheet", "styles.css"));

            string html = doc.ToHtml(options);

            Assert.Contains("<link rel=\"stylesheet\" href=\"styles.css\"", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}
