using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StyleMissing_HandlerProvidesReplacement() {
            string html = "<p class=\"unknown\">Text</p>";
            bool invoked = false;
            var options = new HtmlToWordOptions {
                StyleMissingHandler = e => {
                    invoked = true;
                    e.Style = WordParagraphStyles.Heading1;
                }
            };
            using var doc = html.ToWordDocument(options);

            Assert.True(invoked);
            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
        }

        [Fact]
        public void HtmlToWord_StyleMissing_CreateStyleOnDemand() {
            string html = "<p class=\"custom\">Text</p>";
            string styleId = "DynamicStyle";
            var options = new HtmlToWordOptions {
                StyleMissingHandler = e => {
                    if (e.ClassName == "custom") {
                        WordParagraphStyle.RegisterFontStyle(styleId, "Courier New");
                        e.StyleId = styleId;
                    }
                }
            };
            using var doc = html.ToWordDocument(options);

            Assert.Equal(styleId, doc.Paragraphs[0].StyleId);
        }
    }
}
