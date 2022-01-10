using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTOC() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOC.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.TableOfContent == null, "TableOfContent Should not be set");
                Assert.True(document.Settings.UpdateFieldsOnOpen == false, "UpdateFieldsOnOpen should not be set");

                WordTableOfContent wordTableContent = document.AddTableOfContent(TableOfContentStyle.Template1);
                wordTableContent.Text = "This is Table of Contents";
                wordTableContent.TextNoContent = "Ooopsi, no content";

                document.Settings.UpdateFieldsOnOpen = true;

                document.InsertPageBreak();

                var paragraph = document.InsertParagraph("Test Heading 1");
                paragraph.Heading = "Heading1";

                Assert.True(document.Settings.UpdateFieldsOnOpen == true, "UpdateFieldsOnOpen should be set");
                Assert.True(document.TableOfContent != null, "TableOfContent Should be set");
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.PageBreaks.Count == 1, "PageBreak count should be 1");
                Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 2, "Number of paragraphs on 1st section is wrong.");

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOC.docx"))) {
                Assert.True(document.Settings.UpdateFieldsOnOpen == true, "UpdateFieldsOnOpen should be set");
                Assert.True(document.TableOfContent != null, "TableOfContent Should be set");
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.PageBreaks.Count == 1, "PageBreak count should be 1");
                Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 2, "Number of paragraphs on 1st section is wrong.");

                document.TableOfContent.Text = "This is a test";
                document.TableOfContent.TextNoContent = "This is sub test";
                document.Settings.UpdateFieldsOnOpen = false;

                Assert.True(document.TableOfContent.Text == "This is a test", "TableOfContent Text should be set");
                Assert.True(document.TableOfContent.TextNoContent == "This is sub test", "TableOfContent Text should be set");
                Assert.True(document.Settings.UpdateFieldsOnOpen == false, "UpdateFieldsOnOpen should be set");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOC.docx"))) {
                Assert.True(document.Settings.UpdateFieldsOnOpen == false, "UpdateFieldsOnOpen should not be set");
                Assert.True(document.TableOfContent != null, "TableOfContent Should be set");
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.PageBreaks.Count == 1, "PageBreak count should be 1");
                Assert.True(document.Paragraphs.Count == 2, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 2, "Number of paragraphs on 1st section is wrong.");

                Assert.True(document.TableOfContent.Text == "This is a test", "TableOfContent Text should be set");
                Assert.True(document.TableOfContent.TextNoContent == "This is sub test", "TableOfContent Text should be set");
                document.Save();
            }
        }
    }
}
