using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
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
                new Run(new RunProperties(new Bold()), new Text("link-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new RunProperties(new Italic()), new Text(" link-suffix"))) {
                Anchor = "target"
            });

            string html = document.ToHtml();

            int prefix = html.IndexOf("link-prefix ", StringComparison.Ordinal);
            int math = html.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int suffix = html.IndexOf(" link-suffix", StringComparison.Ordinal);
            Assert.True(prefix >= 0 && prefix < math && math < suffix, html);
            Assert.Contains("aria-label=\"linked\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<strong>link-prefix </strong>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<em> link-suffix</em>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, html.Split(new[] { "link-prefix " }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, html.Split(new[] { " link-suffix" }, StringSplitOptions.None).Length - 1);
            IElement anchor = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(html).QuerySelector("a[href='#target']"));
            Assert.NotNull(anchor.QuerySelector("math"));
            Assert.Contains("link-prefix", anchor.TextContent, StringComparison.Ordinal);
            Assert.Contains("link-suffix", anchor.TextContent, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToHtml_ExportsEquationOnlyHyperlinkAsLinkedMathMl() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new Hyperlink(
                new M.OfficeMath(new M.Run(new M.Text("linked-only")))) {
                Anchor = "target"
            });

            string html = document.ToHtml();

            IElement anchor = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(html).QuerySelector("a[href='#target']"));
            IElement math = Assert.IsAssignableFrom<IElement>(anchor.QuerySelector("math"));
            Assert.Equal("linked-only", math.GetAttribute("aria-label"));
        }

        [Fact]
        public void WordToHtml_PreservesBreakSharingAHyperlinkWithEquation() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new Hyperlink(
                new Run(new Text("prefix")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new Break()),
                new Run(new Text("suffix"))) {
                Anchor = "target"
            });

            string html = document.ToHtml();

            IElement anchor = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(html).QuerySelector("a[href='#target']"));
            Assert.NotNull(anchor.QuerySelector("math"));
            Assert.NotNull(anchor.QuerySelector("br"));
            int math = anchor.InnerHtml.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int lineBreak = anchor.InnerHtml.IndexOf("<br", StringComparison.OrdinalIgnoreCase);
            int suffix = anchor.InnerHtml.IndexOf("suffix", StringComparison.Ordinal);
            Assert.True(math >= 0 && math < lineBreak && lineBreak < suffix, anchor.InnerHtml);
        }

        [Fact]
        public void WordToHtml_PreservesFormControlSharingInlineContainerWithEquation() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(new SdtRun(
                new SdtProperties(
                    new SdtAlias { Val = "Equation approval" },
                    new Tag { Val = "EquationApproval" },
                    new W14.SdtContentCheckBox(new W14.Checked { Val = W14.OnOffValues.One })),
                new SdtContentRun(
                    new Run(new Text("☑")),
                    new M.OfficeMath(new M.Run(new M.Text("approved"))))));

            string html = document.ToHtml();
            IDocument parsed = HtmlDocumentParser.ParseDocument(html);

            IElement input = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("input[type='checkbox'][data-tag='EquationApproval']"));
            IElement math = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("math[aria-label='approved']"));
            Assert.True(input.HasAttribute("checked"), html);
            Assert.Same(input.ParentElement, math.ParentElement);
            Assert.True(
                Array.IndexOf(input.ParentElement!.Children.ToArray(), input) < Array.IndexOf(math.ParentElement!.Children.ToArray(), math),
                input.ParentElement.InnerHtml);
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

        [Fact]
        public void HtmlToWord_PreservesMathMlAndRunFormattingInsideHyperlink() {
            const string html = "<p><a href=\"#target\"><strong>prefix </strong><math aria-label=\"linked\"><mi>x</mi></math><em> suffix</em></a></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            WordParagraph paragraph = Assert.Single(document.Paragraphs);
            Hyperlink hyperlink = Assert.Single(paragraph._paragraph.Elements<Hyperlink>());
            Assert.Equal("target", hyperlink.Anchor?.Value);
            Assert.Collection(
                hyperlink.ChildElements,
                child => Assert.NotNull(Assert.IsType<Run>(child).RunProperties?.Bold),
                child => Assert.IsType<M.OfficeMath>(child),
                child => Assert.NotNull(Assert.IsType<Run>(child).RunProperties?.Italic));
            Assert.Equal("prefix linked suffix", hyperlink.InnerText);
            Assert.Single(document.Equations);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void HtmlToWord_PreservesEquationOnlyMathMlInsideHyperlink() {
            const string html = "<p><a href=\"https://example.test/math\"><math aria-label=\"linked-only\"><mi>x</mi></math></a></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            WordParagraph paragraph = Assert.Single(document.Paragraphs);
            Hyperlink hyperlink = Assert.Single(paragraph._paragraph.Elements<Hyperlink>());
            Assert.Single(hyperlink.Elements<M.OfficeMath>());
            Assert.Empty(hyperlink.Elements<Run>());
            Assert.Equal("https://example.test/math", new WordHyperLink(document, paragraph._paragraph, hyperlink).Uri?.ToString());
            Assert.Single(document.Equations);
            Assert.Empty(document.ValidateDocument());
        }
    }
}
