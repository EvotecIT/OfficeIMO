using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_OpeningWordWithSections() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "BasicDocumentWithSections.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 3);

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 4);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordEmptyDocument() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "EmptyDocument.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1);

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 1);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordEmptyDocumentWithSectionBreak() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "EmptyDocumentWithSection.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs is wrong.");

                // There is only one PageBreak in this document.
                Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }
        [Fact]
        public void Test_OpeningWordDocumentWithSectionBreak() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithSection.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs is wrong.");
                Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");
                
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs in first section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of paragraphs in second section is wrong. Current: " + document.Sections[1].Paragraphs.Count);
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithSectionBreak() {
            //using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithSection.docx"))) {
            //    // There is only one Paragraph at the document level.
            //    Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs is wrong.");

            //    // There is only one PageBreak in this document.
            //    Assert.True(document.Sections.Count == 2, "Number of sections is wrong.");

            //    Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of sections is wrong.");
            //    Assert.True(document.Sections[1].Paragraphs.Count == 1, "Number of sections is wrong.");
            //}
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.InsertParagraph("Test 1");

                var section1 = document.InsertSection();
                section1.InsertParagraph("Test 1");

                document.InsertParagraph("Test 2");
                var section2 = document.InsertSection();

                document.InsertParagraph("Test 3");
                var section3 = document.InsertSection();

                Assert.True(document.Paragraphs.Count == 4, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 2, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 0, "Number of paragraphs on 4th section is wrong.");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSections.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 4, "Number of paragraphs during load is wrong.");
                Assert.True(document.Sections.Count == 4, "Number of sections during load is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[1].Paragraphs.Count == 2, "Number of paragraphs on 2nd section is wrong.");
                Assert.True(document.Sections[2].Paragraphs.Count == 1, "Number of paragraphs on 3rd section is wrong.");
                Assert.True(document.Sections[3].Paragraphs.Count == 0, "Number of paragraphs on 4th section is wrong.");
            }
        }
    }
}
