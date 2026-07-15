using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using M = DocumentFormat.OpenXml.Math;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void WordToHtml_ExportsStructuredEquationsAsMathMl() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            string html = document.ToHtml();

            Assert.Contains("<math", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<mfrac", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"(a)/(b)\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_PreservesEquationOrderWithinParagraph() {
            const string omml = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph.AddEquation(omml);
            paragraph.AddText(" after");

            string html = document.ToHtml();

            int before = html.IndexOf("before ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int after = html.IndexOf(" after", StringComparison.Ordinal);
            Assert.True(before >= 0 && before < math && math < after, html);
        }

        [Fact]
        public void WordToHtml_ExportsEquationInsideVisibleRevisionWrapper() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(new InsertedRun(new M.OfficeMath(new M.Run(new M.Text("tracked")))) {
                Id = "1",
                Author = "Reviewer"
            });
            paragraph.AddText(" after");

            string html = document.ToHtml();

            int before = html.IndexOf("before ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int after = html.IndexOf(" after", StringComparison.Ordinal);
            Assert.True(before >= 0 && before < math && math < after, html);
            Assert.Contains("aria-label=\"tracked\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_ExportsEquationInsideInlineContentControl() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(new SdtRun(
                new SdtProperties(new SdtId { Val = 2076 }),
                new SdtContentRun(
                    new Run(new Text("control-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("controlled"))),
                    new Run(new Text(" control-suffix")))));
            paragraph.AddText(" after");

            string html = document.ToHtml();

            int before = html.IndexOf("before ", StringComparison.Ordinal);
            int prefix = html.IndexOf("control-prefix ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int suffix = html.IndexOf(" control-suffix", StringComparison.Ordinal);
            int after = html.IndexOf(" after", StringComparison.Ordinal);
            Assert.True(before >= 0 && before < prefix && prefix < math && math < suffix && suffix < after, html);
            Assert.Contains("aria-label=\"controlled\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_ExportsEquationAndSurroundingTextInsideHyperlink() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new Hyperlink(
                new Run(new Text("link-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new Text(" link-suffix"))) {
                Anchor = "target"
            });

            string html = document.ToHtml();

            int prefix = html.IndexOf("link-prefix ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int suffix = html.IndexOf(" link-suffix", StringComparison.Ordinal);
            Assert.True(prefix >= 0 && prefix < math && math < suffix, html);
            Assert.Contains("aria-label=\"linked\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_ExportsComplexEqFieldsAsMathMlWithoutCachedTextDuplication() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("(a)/(b)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            paragraph.AddText(" after");

            string html = document.ToHtml();

            int before = html.IndexOf("before ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int after = html.IndexOf(" after", StringComparison.Ordinal);
            Assert.True(before >= 0 && before < math && math < after, html);
            Assert.Contains("<mtext>(a)/(b)</mtext>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(2, html.Split(new[] { "(a)/(b)" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void HtmlToWord_ImportsMathMlAsEditableOmmlTextFallback() {
            const string html = "<p>Formula: <math aria-label=\"sqrt(x)\"><msqrt><mi>x</mi></msqrt></math></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            WordEquation equation = Assert.Single(document.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal("sqrt(x)", equation.Text);
        }
    }
}
