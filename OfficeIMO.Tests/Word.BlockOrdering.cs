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
