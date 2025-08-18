using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text.RegularExpressions;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_CodeBlock_RoundTrip() {
            string html = "<pre><code>var x = 1;\nvar y = 2;</code></pre>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var codeParas = doc.Paragraphs.Where(p => p.StyleId == "HTMLPreformatted" && !string.IsNullOrEmpty(p.Text)).ToList();
            Assert.Equal(2, codeParas.Count);
            Assert.Equal("var x = 1;", codeParas[0].Text);
            Assert.Equal("var y = 2;", codeParas[1].Text);
            foreach (var p in codeParas) {
                Assert.Equal(FontResolver.Resolve("monospace"), p.FontFamily);
            }

            string roundTrip = doc.ToHtml();
            Assert.Contains("<pre><code>", roundTrip);
            Assert.Contains("var x = 1;", roundTrip);
            Assert.Contains("var y = 2;", roundTrip);
            Assert.Single(Regex.Matches(roundTrip, "<pre>"));
        }

        [Fact]
        public void WordToHtml_MonospaceParagraph_OutputCodeBlock() {
            using var document = WordDocument.Create();
            var mono = FontResolver.Resolve("monospace")!;
            document.AddParagraph("Console.WriteLine(\"Hello\");").SetFontFamily(mono).SetStyleId("HTMLPreformatted");

            string html = document.ToHtml();

            Assert.Contains("<pre><code>Console.WriteLine(\"Hello\");</code></pre>", html);
        }
    }
}
