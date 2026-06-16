using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SectionBlockInsertionPreservesCallOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionBlockInsertionOrder.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSection section = document.Sections[0];

                section.AddParagraph("Before table");
                section.AddTable(1, 1);
                section.AddParagraph("After table");
                section.AddTable(1, 1);
                section.AddParagraph("Tail");

                Assert.Equal(
                    new[] {
                        "Paragraph:Before table",
                        "Table",
                        "Paragraph:After table",
                        "Table",
                        "Paragraph:Tail"
                    },
                    GetTopLevelContentOrder(document));
                AssertFinalSectionPropertiesRemainLast(document);
            }
        }

        [Fact]
        public void Test_BodyBlockInsertionPreservesOrderAcrossTableOfContents() {
            string filePath = Path.Combine(_directoryWithFiles, "BodyBlockInsertionOrderWithToc.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro");
                document.AddTable(1, 1);
                document.AddTableOfContent();
                document.AddParagraph("After TOC");

                Assert.Equal(
                    new[] {
                        "Paragraph:Intro",
                        "Table",
                        "SdtBlock",
                        "Paragraph:After TOC"
                    },
                    GetTopLevelContentOrder(document));
                AssertFinalSectionPropertiesRemainLast(document);
            }
        }

        [Fact]
        public void Test_SectionBlockInsertionPreservesOrderAfterDocumentLevelToc() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionBlockInsertionOrderWithToc.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSection section = document.Sections[0];

                section.AddParagraph("Intro");
                section.AddTable(1, 1);
                document.AddTableOfContent();
                section.AddParagraph("After TOC");

                Assert.Equal(
                    new[] {
                        "Paragraph:Intro",
                        "Table",
                        "SdtBlock",
                        "Paragraph:After TOC"
                    },
                    GetTopLevelContentOrder(document));
                AssertFinalSectionPropertiesRemainLast(document);
            }
        }

        [Fact]
        public void Test_SectionBlockInsertionUsesCurrentSectionBoundary() {
            string filePath = Path.Combine(_directoryWithFiles, "EarlierSectionBlockInsertionOrder.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddSection();

                document.Sections[0].AddParagraph("Late first section");
                document.Sections[1].AddParagraph("Second section");
                document.Sections[1].AddTable(1, 1);
                document.Sections[1].AddParagraph("Second section tail");

                Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
                List<OpenXmlElement> children = body.ChildElements.ToList();
                int insertedIndex = children.FindIndex(element => element is Paragraph paragraph && paragraph.InnerText == "Late first section");
                int firstSectionBoundaryIndex = children.FindIndex(element => element is Paragraph paragraph
                    && paragraph.ParagraphProperties?.SectionProperties != null);
                int secondSectionParagraphIndex = children.FindIndex(element => element is Paragraph paragraph && paragraph.InnerText == "Second section");
                int secondSectionTableIndex = children.FindIndex(element => element is Table);
                int secondSectionTailIndex = children.FindIndex(element => element is Paragraph paragraph && paragraph.InnerText == "Second section tail");

                Assert.True(insertedIndex >= 0, "Inserted first-section paragraph should exist.");
                Assert.True(firstSectionBoundaryIndex >= 0, "First section boundary paragraph should exist.");
                Assert.True(secondSectionParagraphIndex >= 0, "Inserted second-section paragraph should exist.");
                Assert.True(secondSectionTableIndex >= 0, "Inserted second-section table should exist.");
                Assert.True(secondSectionTailIndex >= 0, "Inserted second-section tail paragraph should exist.");
                Assert.True(insertedIndex < firstSectionBoundaryIndex, "Blocks appended to an earlier section must stay before that section's boundary.");
                Assert.True(firstSectionBoundaryIndex < secondSectionParagraphIndex, "Blocks appended to the new section must stay after the previous section boundary.");
                Assert.True(secondSectionParagraphIndex < secondSectionTableIndex, "Second-section blocks must preserve insertion order.");
                Assert.True(secondSectionTableIndex < secondSectionTailIndex, "Second-section blocks must preserve insertion order.");
                Assert.DoesNotContain(document.Sections[1].Paragraphs, paragraph => paragraph.Text == "Late first section");
                Assert.DoesNotContain(document.Sections[0].Paragraphs, paragraph => paragraph.Text == "Second section");
                Assert.Contains(document.Sections[1].Paragraphs, paragraph => paragraph.Text == "Second section");
                Assert.Empty(document.Sections[0].Tables);
                Assert.Single(document.Sections[1].Tables);
            }
        }

        [Fact]
        public void Test_RegeneratingTableOfContentsPreservesOriginalBlockPosition() {
            string filePath = Path.Combine(_directoryWithFiles, "RegenerateTocPreservesBlockOrder.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro");
                document.AddTableOfContent();
                document.AddParagraph("After TOC");

                document.RegenerateTableOfContent();

                Assert.Equal(
                    new[] {
                        "Paragraph:Intro",
                        "SdtBlock",
                        "Paragraph:After TOC"
                    },
                    GetTopLevelContentOrder(document));
                AssertFinalSectionPropertiesRemainLast(document);
            }
        }

        [Fact]
        public void Test_BodyBlockInsertionPreservesRawInlineAppendOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "BodyBlockInsertionAfterRawAppend.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro");
                document._document.Body!.Append(new Paragraph(new Run(new Text("Raw inline edit"))));
                document.Sections[0].AddTable(1, 1);
                document.AddParagraph("After raw edit");

                Assert.Equal(
                    new[] {
                        "Paragraph:Intro",
                        "Paragraph:Raw inline edit",
                        "Table",
                        "Paragraph:After raw edit"
                    },
                    GetTopLevelContentOrder(document));
                AssertFinalSectionPropertiesRemainLast(document);
            }
        }

        private static IReadOnlyList<string> GetTopLevelContentOrder(WordDocument document) {
            Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
            return body.ChildElements
                .Where(element => element is Paragraph or Table or SdtBlock)
                .Select(DescribeBlock)
                .ToList();
        }

        private static string DescribeBlock(OpenXmlElement element) {
            return element switch {
                Paragraph paragraph => "Paragraph:" + paragraph.InnerText,
                Table => "Table",
                SdtBlock => "SdtBlock",
                _ => element.GetType().Name
            };
        }

        private static void AssertFinalSectionPropertiesRemainLast(WordDocument document) {
            Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
            SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();

            if (sectionProperties != null) {
                Assert.Same(sectionProperties, body.ChildElements.Last());
            }
        }
    }
}
