using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTabs() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTabs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph1 = document.AddParagraph("Some text before adding tab").AddTab().AddTab().AddText("Test");

                Assert.True(document.ParagraphsTabStops.Count == 0);
                Assert.True(document.ParagraphsTabs.Count == 2);
                Assert.True(document.Paragraphs[0].Text == "Some text before adding tab");
                Assert.True(document.Paragraphs.Count == 4);

                Assert.True(document.Paragraphs[1].IsTab == true);
                Assert.True(document.Paragraphs[2].IsTab == true);

                Assert.True(paragraph1.IsTab == false);

                var paragraph2 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER").AddTab();

                Assert.True(document.Paragraphs.Count == 6);
                Assert.True(paragraph2.IsTab == true);

                Assert.True(document.ParagraphsTabs.Count == 3);

                paragraph2.Tab.Remove();

                Assert.True(document.Paragraphs.Count == 5);

                Assert.True(document.ParagraphsTabs.Count == 2);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTabs.docx"))) {
                Assert.True(document.ParagraphsTabStops.Count == 0);
                Assert.True(document.ParagraphsTabs.Count == 2);
                Assert.True(document.Paragraphs[1].IsTab == true);
                Assert.True(document.Paragraphs[2].IsTab == true);
                Assert.True(document.Sections[0].ParagraphsTabStops.Count == 0);
                Assert.True(document.Sections[0].ParagraphsTabs.Count == 2);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTabs.docx"))) {


            }
        }
    }
}
