using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_CodeBlock_RoundTrip() {
            string html = "<pre><code>var x = 1;\nvar y = 2;</code></pre>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

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
        public void HtmlToWord_PreInTableCell_PreservesAllLines() {
            string html = "<table><tr><td><pre>a\nb</pre></td></tr></table>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var paragraphs = doc.Tables[0].Rows[0].Cells[0].Paragraphs
                .Where(paragraph => paragraph.StyleId == "HTMLPreformatted")
                .Select(paragraph => paragraph.Text)
                .Where(text => !string.IsNullOrEmpty(text))
                .ToArray();
            Assert.Equal(new[] { "a", "b" }, paragraphs);
        }

        [Fact]
        public void HtmlToWord_InlineCode_StaysInParagraphAndRoundTripsAsCode() {
            string html = "<p>Use <code>dotnet test</code> now.</p>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var bodyParagraphs = doc._wordprocessingDocument!.MainDocumentPart!.Document!.Body!
                .Elements<Paragraph>()
                .Where(p => !string.IsNullOrEmpty(p.InnerText))
                .ToArray();
            Assert.Single(bodyParagraphs);
            Assert.Equal("Use dotnet test now.", bodyParagraphs[0].InnerText);

            var runs = bodyParagraphs[0].Elements<Run>().Where(r => !string.IsNullOrEmpty(r.InnerText)).ToArray();
            Assert.Equal(new[] { "Use ", "dotnet test", " now." }, runs.Select(r => r.InnerText).ToArray());
            Assert.Equal(FontResolver.Resolve("monospace"), runs[1].RunProperties!.RunFonts!.Ascii!.Value);
            Assert.Equal("HtmlCode", runs[1].RunProperties!.RunStyle!.Val!.Value);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<p>Use <code>dotnet test</code> now.</p>", roundTrip);
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
