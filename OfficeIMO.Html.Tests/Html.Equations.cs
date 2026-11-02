using AngleSharp.Dom;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void WordToHtml_ExportedStructuredEquationsSurviveSavedWordRoundTrip() {
            const string omml = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
                + "<m:f><m:num><m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e>"
                + "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:num>"
                + "<m:den><m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg>"
                + "<m:e><m:r><m:t>b</m:t></m:r></m:e></m:rad></m:den></m:f></m:oMath>";
            using WordDocument source = WordDocument.Create();
            WordParagraph paragraph = source.AddParagraph("before ");
            paragraph.AddEquation(omml);
            paragraph.AddText(" after");
            string expectedText = Assert.Single(source.Equations).Text;
            Assert.Empty(source.ValidateDocument());

            HtmlTextConversionResult export = source.ConvertToHtmlWithEvidence();
            string html = export.RequireValue();
            IElement exportedMath = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(html).QuerySelector("math"));
            Assert.NotNull(exportedMath.QuerySelector("mfrac"));
            Assert.NotNull(exportedMath.QuerySelector("msup"));
            Assert.NotNull(exportedMath.QuerySelector("mroot"));
            Assert.False(string.IsNullOrWhiteSpace(exportedMath.GetAttribute("aria-label")));

            HtmlToWordResult import = HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence();
            using WordDocument imported = import.RequireValue();
            using var artifact = new MemoryStream();
            imported.Save(artifact);
            byte[] bytes = artifact.ToArray();
            using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes),
                new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            WordEquation equation = Assert.Single(reopened.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal(expectedText, equation.Text);
            Assert.Empty(reopened.ValidateDocument());

            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XElement native = ReadSavedHtmlEquation(bytes);
            XElement fraction = Assert.Single(native.Descendants(math + "f"));
            XElement script = Assert.Single(fraction.Element(math + "num")!.Elements(math + "sSup"));
            Assert.Equal("a", script.Element(math + "e")!.Value);
            Assert.Equal("2", script.Element(math + "sup")!.Value);
            XElement radical = Assert.Single(fraction.Element(math + "den")!.Elements(math + "rad"));
            Assert.Equal("3", radical.Element(math + "deg")!.Value);
            Assert.Equal("b", radical.Element(math + "e")!.Value);
            XElement savedParagraph = Assert.Single(native.Ancestors(word + "p"));
            Assert.Equal("before {equation} after", string.Concat(savedParagraph.Elements().Select(element =>
                element.Name == math + "oMath" ? "{equation}"
                    : string.Concat(element.Descendants(word + "t").Select(text => text.Value)))));
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
                    new Run(new Text("control-prefix ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                    new M.OfficeMath(new M.Run(new M.Text("controlled"))),
                    new Run(new Text(" control-suffix") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))));
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
                new Run(new RunProperties(new Bold()), new Text("link-prefix ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new RunProperties(new Italic()), new Text(" link-suffix") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) {
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
                    new M.OfficeMath(new M.Run(new M.Text("approved"))),
                    new Run(new Text("☑")))));

            string html = document.ToHtml();
            IDocument parsed = HtmlDocumentParser.ParseDocument(html);

            IElement input = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("input[type='checkbox'][data-tag='EquationApproval']"));
            IElement math = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("math[aria-label='approved']"));
            Assert.True(input.HasAttribute("checked"), html);
            Assert.Same(input.ParentElement, math.ParentElement);
            Assert.True(
                Array.IndexOf(math.ParentElement!.Children.ToArray(), math) < Array.IndexOf(input.ParentElement!.Children.ToArray(), input),
                input.ParentElement.InnerHtml);
        }

        [Fact]
        public void WordToHtml_PreservesPictureControlSharingInlineContainerWithEquation() {
            string imagePath = Path.Combine(RepositoryTestPaths.Find(), "Assets", "OfficeIMO.png");
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            WordPictureControl picture = paragraph.AddPictureControl(imagePath, alias: "Equation picture", tag: "EquationPicture");
            picture._sdtRun.SdtContentRun!.Append(
                new M.OfficeMath(new M.Run(new M.Text("pictured-equation"))));

            IDocument parsed = HtmlDocumentParser.ParseDocument(document.ToHtml());
            IElement image = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("img"));
            IElement math = Assert.IsAssignableFrom<IElement>(parsed.QuerySelector("math[aria-label='pictured-equation']"));

            Assert.StartsWith("data:image/png;base64,", image.GetAttribute("src"), StringComparison.Ordinal);
            Assert.Same(image.ParentElement, math.ParentElement);
            Assert.True(
                Array.IndexOf(image.ParentElement!.Children.ToArray(), image) < Array.IndexOf(math.ParentElement!.Children.ToArray(), math),
                image.ParentElement.InnerHtml);
        }

        [Fact]
        public void WordToHtml_ExpandsHyperlinkWhoseVisibleEquationContentIsWrappedByInlineControl() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(new Hyperlink(
                new SdtRun(
                    new SdtProperties(new SdtId { Val = 2078 }),
                    new SdtContentRun(
                        new Run(new Text("nested-prefix ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                        new M.OfficeMath(new M.Run(new M.Text("nested-linked"))),
                        new Run(new Text(" nested-suffix") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })))) {
                Anchor = "nested-equation-target"
            });
            paragraph.AddText(" after");

            string html = document.ToHtml();

            IElement anchor = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(html).QuerySelector("a[href='#nested-equation-target']"));
            Assert.Equal("nested-prefix nested-linked nested-suffix", anchor.TextContent);
            Assert.NotNull(anchor.QuerySelector("math[aria-label='nested-linked']"));
            Assert.Equal(1, html.Split(new[] { "nested-prefix" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, html.Split(new[] { "nested-suffix" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void WordToHtml_PreservesInnerHyperlinkOnEquationInsideInlineControl() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new SdtRun(
                new SdtProperties(new SdtId { Val = 2079 }),
                new SdtContentRun(
                    new Hyperlink(
                        new Run(new Text("inner-prefix ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                        new M.OfficeMath(new M.Run(new M.Text("inner-linked"))),
                        new Run(new Text(" inner-suffix") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) {
                        Anchor = "inner-equation-target"
                    })));

            IDocument parsed = HtmlDocumentParser.ParseDocument(document.ToHtml());
            IElement linkedMath = Assert.IsAssignableFrom<IElement>(
                parsed.QuerySelector("a[href='#inner-equation-target'] math[aria-label='inner-linked']"));

            Assert.Equal("a", linkedMath.ParentElement?.LocalName);
            Assert.Equal("#inner-equation-target", linkedMath.ParentElement?.GetAttribute("href"));
        }

        [Fact]
        public void WordToHtml_PreservesSemanticRunTagsBesideEquation() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new Hyperlink(
                new Run(new RunProperties(new RunStyle { Val = "HtmlCode" }), new Text("code")),
                new Run(new RunProperties(new RunStyle { Val = "HtmlDeletedText" }, new Strike()), new Text("deleted")),
                new M.OfficeMath(new M.Run(new M.Text("semantic-equation"))),
                new Run(new RunProperties(new RunStyle { Val = "HtmlInsertedText" }, new Underline { Val = UnderlineValues.Single }), new Text("inserted")),
                new Run(new RunProperties(new RunStyle { Val = "HtmlMarkedText" }, new Highlight { Val = HighlightColorValues.Yellow }), new Text("marked")),
                new Run(new RunProperties(new RunStyle { Val = "HtmlCite" }), new Text("citation")),
                new Run(new RunProperties(new RunStyle { Val = "HtmlDfn" }), new Text("definition"))) {
                Anchor = "semantic-equation-target"
            });

            IElement anchor = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(document.ToHtml()).QuerySelector("a[href='#semantic-equation-target']"));

            Assert.Equal("code", anchor.QuerySelector("code")?.TextContent);
            Assert.Equal("deleted", anchor.QuerySelector("del")?.TextContent);
            Assert.NotNull(anchor.QuerySelector("math[aria-label='semantic-equation']"));
            Assert.Equal("inserted", anchor.QuerySelector("ins")?.TextContent);
            Assert.Equal("marked", anchor.QuerySelector("mark")?.TextContent);
            Assert.Equal("citation", anchor.QuerySelector("cite")?.TextContent);
            Assert.Equal("definition", anchor.QuerySelector("dfn")?.TextContent);
            Assert.Null(anchor.QuerySelector("del s"));
            Assert.Null(anchor.QuerySelector("ins u"));
        }

        [Fact]
        public void WordToHtml_ExpandsVisibleRevisionTextAroundEquation() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            paragraph._paragraph.Append(new InsertedRun(
                new Run(new Text("revision-prefix ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                new M.OfficeMath(new M.Run(new M.Text("revision-equation"))),
                new Run(new Text(" revision-suffix") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) {
                Id = "2078",
                Author = "Reviewer"
            });
            paragraph.AddText(" after");

            IElement output = Assert.IsAssignableFrom<IElement>(
                HtmlDocumentParser.ParseDocument(document.ToHtml()).QuerySelector("p"));
            int prefix = output.InnerHtml.IndexOf("revision-prefix", StringComparison.Ordinal);
            int math = output.InnerHtml.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int suffix = output.InnerHtml.IndexOf("revision-suffix", StringComparison.Ordinal);

            Assert.True(prefix >= 0 && prefix < math && math < suffix, output.InnerHtml);
            Assert.Contains("before revision-prefix revision-equation revision-suffix after", output.TextContent, StringComparison.Ordinal);
            Assert.Equal(1, output.TextContent.Split(new[] { "revision-prefix" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, output.TextContent.Split(new[] { "revision-suffix" }, StringSplitOptions.None).Length - 1);
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

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void HtmlToWord_ImportsStructuredMathMlRadicalWithOptionalLabel(bool labeled) {
            string label = labeled ? " aria-label='spoken radical'" : string.Empty;
            string html = "<p>Formula: <math" + label + "><msqrt><mi>x</mi></msqrt></math></p>";

            using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();

            WordEquation equation = Assert.Single(document.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal("sqrt(x)", equation.Text);
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XElement radical = Assert.Single(XElement.Parse(equation.Omml!).Descendants(math + "rad"));
            Assert.Equal("x", radical.Element(math + "e")!.Value);
            Assert.Empty(document.ValidateDocument());
        }

        [Theory]
        [InlineData("<math aria-label='spoken fallback'></math>", "spoken fallback")]
        [InlineData("<math>x + 1</math>", "x + 1")]
        [InlineData("<math aria-label='annotation fallback'><semantics>"
            + "<annotation encoding='application/x-tex'>ignored TeX</annotation>"
            + "<annotation-xml encoding='application/mathml+xml'><math><mo lspace='.5em'>+</mo></math>"
            + "</annotation-xml></semantics></math>", "annotation fallback")]
        [InlineData("<math><semantics><annotation encoding='application/x-tex'>x + 1</annotation>"
            + "</semantics></math>", "x + 1")]
        public void HtmlToWord_PreservesEquationTextFallbackWhenNoPresentationProjectionExists(
            string html, string expectedText) {
            HtmlToWordResult result = HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence();
            using WordDocument document = result.RequireValue();
            using var artifact = new MemoryStream();
            document.Save(artifact);
            byte[] bytes = artifact.ToArray();
            using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes),
                new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

            WordEquation equation = Assert.Single(reopened.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal(expectedText, equation.Text);
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XElement native = ReadSavedHtmlEquation(bytes);
            Assert.Equal(math + "r", Assert.Single(native.Elements()).Name);
            Assert.Equal(expectedText, Assert.Single(native.Descendants(math + "t")).Value);
            Assert.Empty(reopened.ValidateDocument());
            Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
                diagnostic.Component == "OfficeIMO.Word.Html"
                && diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated);
        }

        [Fact]
        public void HtmlToWord_ImportsUnlabeledStructuredMathMlAsEditableOmml() {
            const string html = "<p>Formula: <math><mfrac><mi>a</mi><mi>b</mi></mfrac></math></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();

            WordEquation equation = Assert.Single(document.Equations);
            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal("\\frac{a}{b}", equation.ToLatex());
            Assert.Contains("<mfrac>", equation.ToMathMl(), StringComparison.Ordinal);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void HtmlToWord_PreservesMathMlAndRunFormattingInsideHyperlink() {
            const string html = "<p><a href=\"#target\"><strong>prefix </strong><math aria-label=\"linked\"><mi>x</mi></math><em> suffix</em></a></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();

            WordParagraph paragraph = Assert.Single(document.Paragraphs);
            Hyperlink hyperlink = Assert.Single(paragraph._paragraph.Elements<Hyperlink>());
            Assert.Equal("target", hyperlink.Anchor?.Value);
            Assert.Collection(
                hyperlink.ChildElements,
                child => Assert.NotNull(Assert.IsType<Run>(child).RunProperties?.Bold),
                child => Assert.IsType<M.OfficeMath>(child),
                child => Assert.NotNull(Assert.IsType<Run>(child).RunProperties?.Italic));
            Assert.Equal("prefix x suffix", hyperlink.InnerText);
            Assert.Equal("x", Assert.Single(document.Equations).Text);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void HtmlToWord_PreservesEquationOnlyMathMlInsideHyperlink() {
            const string html = "<p><a href=\"https://example.test/math\"><math aria-label=\"linked-only\"><mi>x</mi></math></a></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();

            WordParagraph paragraph = Assert.Single(document.Paragraphs);
            Hyperlink hyperlink = Assert.Single(paragraph._paragraph.Elements<Hyperlink>());
            Assert.Single(hyperlink.Elements<M.OfficeMath>());
            Assert.Empty(hyperlink.Elements<Run>());
            Assert.Equal("https://example.test/math", new WordHyperLink(document, paragraph._paragraph, hyperlink).Uri?.ToString());
            Assert.Equal("x", Assert.Single(document.Equations).Text);
            Assert.Empty(document.ValidateDocument());
        }

        private static XElement ReadSavedHtmlEquation(byte[] bytes) {
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            using Stream documentXml = package.MainDocumentPart!.GetStream();
            XDocument native = XDocument.Load(documentXml);
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            return Assert.Single(native.Descendants(math + "oMath"));
        }
    }
}
