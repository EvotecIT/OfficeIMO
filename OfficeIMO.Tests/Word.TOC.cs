using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

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

                Assert.True(document.Settings.UpdateFieldsOnOpen == false);

                wordTableContent.Update();

                Assert.True(document.Settings.UpdateFieldsOnOpen == true);

                document.Settings.UpdateFieldsOnOpen = true;

                document.AddPageBreak();

                var paragraph = document.AddParagraph("Test Heading 1");
                paragraph.Style = WordParagraphStyles.Heading1;

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

        [Fact]
        public void Test_CreatingWordDocumentWithTOCAndList() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOCandList.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Settings.UpdateFieldsOnOpen == false, "Update field settings should be turned off for new document");

                document.Settings.UpdateFieldsOnOpen = true;
                document.AddTableOfContent(tableOfContentStyle: TableOfContentStyle.Template2);
                document.AddHeadersAndFooters();
                //var pageNumber = document.Header.Default.AddPageNumber(WordPageNumberStyle.Circle);
                var pageNumber = document.Footer.Default.AddPageNumber(WordPageNumberStyle.VerticalOutline2);
                pageNumber.ParagraphAlignment = JustificationValues.Center;

                document.AddPageBreak();

                var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

                Assert.True(document.Lists.Count == 1, "Lists count should be 1, just TOC");

                wordListToc.AddItem("This is first item");

                wordListToc.AddItem("This is second item");

                document.AddPageBreak();

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.2", 2);

                var para = document.AddParagraph("Let's show everyone how to create a list within already defined list");
                para.CapsStyle = CapsStyle.Caps;
                para.Highlight = HighlightColorValues.DarkMagenta;

                var wordList = document.AddList(WordListStyle.Bulleted);
                wordList.AddItem("List Item 1");
                wordList.AddItem("List Item 2");
                wordList.AddItem("List Item 3");
                wordList.AddItem("List Item 3.1", 1);
                wordList.AddItem("List Item 3.2", 1);
                wordList.AddItem("List Item 3.3", 2);

                wordListToc.AddItem("Text 2.3", 2);

                wordListToc.AddItem("Text 3.3", 3);

                Assert.True(document.Lists.Count == 2, "Lists count should be 2, just TOC + Bullets");
                Assert.True(document.Settings.UpdateFieldsOnOpen == true, "Update field settings should be turned on when it was enabled");
                Assert.True(document.Paragraphs.Count == 17, "All paragraphs including from lists and toc should be here");
                Assert.True(document.PageBreaks.Count == 2, "All page breaks should be shown");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Lists.Count == 2, "Lists count should be 2, just TOC + Bullets");
                Assert.True(document.Settings.UpdateFieldsOnOpen == true, "Update field settings should be turned on when it was enabled");
                Assert.True(document.Paragraphs.Count == 17, "All paragraphs including from lists and toc should be here");
                Assert.True(document.PageBreaks.Count == 2, "All page breaks should be shown");

                // we loaded document, lets add some text to continue 
                document.AddParagraph().SetColor(Color.CornflowerBlue).SetText("This is some text");

                // we loaded document, lets add page break to continue
                document.AddPageBreak();

                // lets find a list which has items which suggest it's a TOC attached list
                WordList wordListToc = null;
                foreach (var list in document.Lists) {
                    if (list.IsToc) {
                        wordListToc = list;
                    }
                }

                // finally lets add another list item
                if (wordListToc != null) {
                    wordListToc.AddItem("Text 4.4", 2);
                }

                document.Settings.UpdateFieldsOnOpen = true;

                Assert.True(document.Lists.Count == 2, "Lists count should be 2, just TOC + Bullets");
                Assert.True(document.Settings.UpdateFieldsOnOpen == true, "Update field settings should be turned on when it was enabled");
                Assert.True(document.Paragraphs.Count == 20, "All paragraphs including from lists and toc should be here");
                Assert.True(document.PageBreaks.Count == 3, "All page breaks should be shown");

                Assert.True(document.Lists[0].IsToc == true, "This list should be TOC");
                Assert.True(document.Lists[1].IsToc == false, "This list should not be TOC");
                Assert.True(document.Lists[0].ListItems[0].Text == document.Paragraphs[1].Text, "Text should be identical");

                Assert.True(document.Lists[0].ListItems[document.Lists[0].ListItems.Count - 1].Text == document.Paragraphs[document.Paragraphs.Count - 1].Text, "Text should be identical");

                document.Save(false);
            }
        }

    }
}
