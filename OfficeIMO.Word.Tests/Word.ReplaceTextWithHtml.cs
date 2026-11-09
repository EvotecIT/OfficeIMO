using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ReplaceTextWithHtmlFragment_PreservesTextAndPlacementForMatchesInOneRun() {
            using var output = new MemoryStream();
            using (WordDocument document = WordDocument.Create()) {
                WordParagraph original = document.AddParagraph("first TOKEN middle TTOKEN tail");
                original.Bold = true;
                original.PageBreakBeforeOverride = true;
                int count = document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>");

                Assert.Equal(2, count);
                document.Save(output);
            }

            output.Position = 0;
            using WordDocument reopened = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Collection(reopened._document.Body!.ChildElements,
                item => Assert.Equal("first ", Assert.IsType<Paragraph>(item).InnerText),
                item => Assert.IsType<AltChunk>(item),
                item => Assert.Equal(" middle T", Assert.IsType<Paragraph>(item).InnerText),
                item => Assert.IsType<AltChunk>(item),
                item => Assert.Equal(" tail", Assert.IsType<Paragraph>(item).InnerText),
                item => Assert.IsType<SectionProperties>(item));
            Assert.All(reopened.Paragraphs, paragraph => Assert.True(paragraph.Bold));
            Assert.True(reopened.Paragraphs[0].PageBreakBeforeOverride);
            Assert.False(reopened.Paragraphs[1].PageBreakBeforeOverride);
            Assert.False(reopened.Paragraphs[2].PageBreakBeforeOverride);
            Assert.Equal(2, reopened.EmbeddedDocuments.Count);
            Assert.All(reopened.EmbeddedDocuments, fragment => {
                string html = Assert.IsType<string>(fragment.GetHtml());
                Assert.StartsWith("<html>", html);
                Assert.Contains("<meta charset=\"utf-8\">", html);
                Assert.Contains("<p>Injected</p>", html);
                Assert.EndsWith("</body></html>", html);
            });
        }

        [Fact]
        public void ReplaceTextWithHtmlFragment_CultureComparisonPreservesTextWithDifferentMatchWidth() {
            using WordDocument document = WordDocument.Create();
            const string text = "caf\u00e9 next";
            WordParagraph paragraph = document.AddParagraph(text);

            int count = document.ReplaceTextWithHtmlFragment("cafe\u0301", "<html><p>Injected</p></html>",
                stringComparison: StringComparison.InvariantCulture);

            Assert.Equal(0, count);
            Assert.Equal(text, paragraph.Text);
            Assert.Empty(document.EmbeddedDocuments);
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_MultiParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlMulti.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro start");
                document.AddParagraph("finish end");
                var count = document.ReplaceTextWithHtmlFragment("startfinish", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal("Intro ", document.Paragraphs[0].Text);
                Assert.Equal(" end", document.Paragraphs[1].Text);
                var body = document._document.Body!;
                Assert.IsType<Paragraph>(body.ChildElements[0]);
                Assert.IsType<AltChunk>(body.ChildElements[1]);
                Assert.IsType<Paragraph>(body.ChildElements[2]);
                Assert.IsType<SectionProperties>(body.ChildElements[3]);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_BoundaryStart() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlBoundaryStart.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("target middle");
                var count = document.ReplaceTextWithHtmlFragment("target", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal(" middle", document.Paragraphs[0].Text);
                Assert.IsType<AltChunk>(document._document.Body!.FirstChild);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_BoundaryEnd() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlBoundaryEnd.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Some text");
                document.AddParagraph("final target");
                var count = document.ReplaceTextWithHtmlFragment("target", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal("final ", document.Paragraphs[1].Text);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }

        [Fact]
        public void ReplaceTextWithHtmlFragment_SplitsOrdinaryHyperlinkAndKeepsSectionEndWithSuffix() {
            using var output = new MemoryStream();
            using (WordDocument document = WordDocument.Create()) {
                WordParagraph paragraph = document.AddParagraph("before TO");
                paragraph.Bold = true;
                paragraph.AddHyperLink("KEN linked TOKEN rest", new System.Uri("https://example.com"));
                paragraph.AddText(" after").Italic = true;
                paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph._paragraph.ParagraphProperties.SectionProperties =
                    new SectionProperties(new SectionType { Val = SectionMarkValues.Continuous });
                document.AddParagraph("next section");

                Assert.Equal(2, document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>"));
                document.Save(output);
            }

            output.Position = 0;
            using WordDocument reopened = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            var children = reopened._document.Body!.ChildElements;
            Paragraph before = Assert.IsType<Paragraph>(children[0]);
            Assert.IsType<AltChunk>(children[1]);
            Paragraph middle = Assert.IsType<Paragraph>(children[2]);
            Assert.IsType<AltChunk>(children[3]);
            Paragraph after = Assert.IsType<Paragraph>(children[4]);
            Assert.Equal("before ", before.InnerText);
            Assert.Equal(" linked ", middle.InnerText);
            Assert.Equal(" rest after", after.InnerText);
            Assert.NotNull(before.GetFirstChild<Run>()!.RunProperties!.Bold);
            Assert.NotNull(after.GetFirstChild<Run>()!.RunProperties!.Italic);
            Assert.Null(before.ParagraphProperties?.SectionProperties);
            Assert.Null(middle.ParagraphProperties?.SectionProperties);
            Assert.Equal(SectionMarkValues.Continuous, after.ParagraphProperties!.SectionProperties!.GetFirstChild<SectionType>()!.Val!.Value);
            Assert.Equal("next section", Assert.IsType<Paragraph>(children[5]).InnerText);
            Assert.IsType<SectionProperties>(children[6]);

            Hyperlink leftLink = Assert.Single(middle.Elements<Hyperlink>());
            Hyperlink rightLink = Assert.Single(after.Elements<Hyperlink>());
            Assert.Equal(" linked ", leftLink.InnerText);
            Assert.Equal(" rest", rightLink.InnerText);
            Assert.NotEqual(leftLink.Id!.Value, rightLink.Id!.Value);
            var relationships = System.Linq.Enumerable.ToDictionary(
                reopened._document.MainDocumentPart!.HyperlinkRelationships, link => link.Id);
            Assert.Equal(2, relationships.Count);
            Assert.Equal("https://example.com/", relationships[leftLink.Id.Value!].Uri.AbsoluteUri);
            Assert.Equal("https://example.com/", relationships[rightLink.Id.Value!].Uri.AbsoluteUri);
        }

        [Fact]
        public void ReplaceTextWithHtmlFragment_KeepsNestedListIndentationWithoutAddingABullet() {
            using WordDocument document = WordDocument.Create();
            WordParagraph original = document.AddList(WordListStyle.Bulleted).AddItem("before TOKEN after", level: 2);
            original.IndentationAfter = global::OfficeIMO.OfficeLength.Points(12);
            var sourceIndentation = WordDocumentTraversal.GetListInfo(original)?.LeftIndentTwips;
            Assert.True(sourceIndentation > 0);

            Assert.Equal(1, document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>"));

            var paragraphs = document.Paragraphs;
            Assert.Equal(2, paragraphs.Count);
            Assert.True(paragraphs[0].IsListItem);
            Assert.False(paragraphs[1].IsListItem);
            Assert.Equal("before ", paragraphs[0].Text);
            Assert.Equal(" after", paragraphs[1].Text);
            Assert.Equal((long)sourceIndentation!.Value, paragraphs[1].IndentationBefore!.Value.ToTwips());
            Assert.Equal(original.IndentationAfter, paragraphs[1].IndentationAfter);
            Assert.Null(paragraphs[1].IndentationHanging);
        }

        [Theory]
        [InlineData("TOKEN after", "direct")]
        [InlineData("TOKEN", "direct")]
        [InlineData("TOKEN after", "style")]
        [InlineData("TOKEN after", "defaults")]
        [InlineData("KEN after", "boundary")]
        public void ReplaceTextWithHtmlFragment_WouldMovePageBreakLeavesAllMatchesAndPartsIntact(string text, string source) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("before TOKEN after");
            if (source == "boundary") {
                document.AddParagraph("later TO");
            }
            WordParagraph target = document.AddParagraph(text);
            Styles styles = document._document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            switch (source) {
                case "style":
                    styles.Append(new WordParagraphStyleDefinition("FragmentBreakBase") { PageBreakBefore = true }.ToOpenXml());
                    styles.Append(new WordParagraphStyleDefinition("FragmentBreakChild") { BasedOnStyleId = "FragmentBreakBase" }.ToOpenXml());
                    target.SetStyleId("FragmentBreakChild");
                    break;
                case "defaults":
                    DocDefaults defaults = styles.DocDefaults ??= new DocDefaults();
                    ParagraphPropertiesDefault paragraphDefaults = defaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
                    ParagraphPropertiesBaseStyle properties = paragraphDefaults.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
                    properties.AddChild(new PageBreakBefore(), true);
                    break;
                default:
                    target.PageBreakBeforeOverride = true;
                    break;
            }
            string originalBody = document._document.Body!.OuterXml;

            Assert.Throws<NotSupportedException>(() => document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>"));

            Assert.Equal(originalBody, document._document.Body!.OuterXml);
            Assert.Empty(document._document.MainDocumentPart!.AlternativeFormatImportParts);
        }

        [Fact]
        public void ReplaceTextWithHtmlFragment_FieldMatchLeavesPlainMatchAndPartsIntact() {
            using WordDocument document = WordDocument.Create();
            WordParagraph plain = document.AddParagraph("before TOKEN after");
            WordParagraph field = document.AddParagraph();
            var marker = new SimpleField(new Run(new Text("TOKEN"))) { Instruction = "PAGE" };
            field._paragraph.Append(marker);

            Assert.Throws<System.NotSupportedException>(() => document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>"));

            Assert.Equal("before TOKEN after", plain.Text);
            Assert.Equal("TOKEN", marker.InnerText);
            Assert.Empty(document.EmbeddedDocuments);
            Assert.Empty(document._document.MainDocumentPart!.AlternativeFormatImportParts);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ReplaceTextWithHtmlFragment_ParagraphRevisionLeavesAllMatchesAndPartsIntact(bool revisionOnEmptyMiddleParagraph) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("before TOKEN after");
            if (revisionOnEmptyMiddleParagraph) {
                document.AddParagraph("later TO");
                WordParagraph revised = document.AddParagraph();
                revised._paragraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphMarkRunProperties(
                        new ParagraphMarkRunPropertiesChange(new PreviousParagraphMarkRunProperties(new Bold())) {
                            Author = "Reviewer", Id = "1"
                        }));
                document.AddParagraph("KEN after");
            } else {
                WordParagraph revised = document.AddParagraph("later TOKEN after");
                revised._paragraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphPropertiesChange(new PreviousParagraphProperties()) {
                        Author = "Reviewer", Id = "1"
                    });
            }
            string originalBody = document._document.Body!.OuterXml;

            Assert.Throws<NotSupportedException>(() => document.ReplaceTextWithHtmlFragment("TOKEN", "<p>Injected</p>"));

            Assert.Equal(originalBody, document._document.Body!.OuterXml);
            Assert.Empty(document._document.MainDocumentPart!.AlternativeFormatImportParts);
        }
    }
}
