using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StyleMissing_EventProvidesReplacement() {
            string html = "<p class=\"unknown\">Text</p>";
            bool invoked = false;
            EventHandler<StyleMissingEventArgs> handler = (s, e) => {
                invoked = true;
                e.Style = WordParagraphStyles.Heading1;
            };
            WordHtmlConverterExtensions.StyleMissing += handler;
            try {
                var doc = html.LoadFromHtml();
                Assert.True(invoked);
                Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
            } finally {
                WordHtmlConverterExtensions.StyleMissing -= handler;
            }
        }

        [Fact]
        public void HtmlToWord_StyleMissing_CreateStyleOnDemand() {
            string html = "<p class=\"custom\">Text</p>";
            string styleId = "DynamicStyle";
            EventHandler<StyleMissingEventArgs> handler = (s, e) => {
                if (e.ClassName == "custom") {
                    WordParagraphStyle.RegisterFontStyle(styleId, "Courier New");
                    e.StyleId = styleId;
                }
            };
            WordHtmlConverterExtensions.StyleMissing += handler;
            try {
                var doc = html.LoadFromHtml();
                Assert.Equal(styleId, doc.Paragraphs[0].StyleId);
            } finally {
                WordHtmlConverterExtensions.StyleMissing -= handler;
            }
        }
    }
}
