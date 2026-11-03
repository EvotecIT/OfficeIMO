using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ComplexSelector_Specificity() {
            string html = "<style>.highlight { color:#0000ff; } div .highlight { color:#ff0000; }</style><div><p class=\"highlight\">Test</p></div>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new HtmlWordExportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal(OfficeColor.Parse("FF0000"), run.Color);
        }

        [Fact]
        public void HtmlToWord_InvalidCss_Ignored() {
            string html = "<style>p { color:#00ff00 } .invalid {</style><p>Test</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new HtmlWordExportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal(OfficeColor.Parse("00FF00"), run.Color);
        }
    }
}
