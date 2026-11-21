using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StyleMissing_HandlerProvidesReplacement() {
            string html = "<p class=\"unknown\">Text</p>";
            bool invoked = false;
            var options = new HtmlWordExportOptions {
                StyleMissingHandler = e => {
                    invoked = true;
                    e.Style = WordParagraphStyles.Heading1;
                }
            };
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(options);

            Assert.True(invoked);
            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
        }

        [Fact]
        public void HtmlToWord_StyleMissing_CreateStyleOnDemand() {
            string html = "<p class=\"custom\">Text</p>";
            string styleId = "DynamicStyle";
            using var doc = WordDocument.Create();
            var options = new HtmlWordExportOptions {
                StyleMissingHandler = e => {
                    if (e.ClassName == "custom") {
                        doc.Styles.AddParagraphStyle(styleId).FontFamily = "Courier New";
                        e.StyleId = styleId;
                    }
                }
            };
            doc.AddHtml(html, options);

            Assert.Equal(styleId, doc.Paragraphs[0].StyleId);
            Assert.Equal("Courier New", doc.Styles[styleId].FontFamily);
        }
    }
}
