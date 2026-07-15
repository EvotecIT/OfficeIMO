using OfficeIMO.Word;
using OfficeIMO.Word.Html;
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
        public void HtmlToWord_ImportsMathMlAsEditableOmmlTextFallback() {
            const string html = "<p>Formula: <math aria-label=\"sqrt(x)\"><msqrt><mi>x</mi></msqrt></math></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            WordEquation equation = Assert.Single(document.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal("sqrt(x)", equation.Text);
        }
    }
}
