using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithLists() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("First List");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                Assert.True(document.Paragraphs[0].IsEmpty == false, "Paragraph is empty");

                Assert.True(document.Lists.Count == 0, "List count matches");

                WordList wordList1 = document.AddList(WordListStyle.Heading1ai);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                Assert.True(document.Paragraphs[0].IsListItem == false, "Paragraph is empty");
                Assert.True(document.Paragraphs[1].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[2].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[3].IsListItem == true, "Paragraph is list item 2");
                Assert.True(document.Paragraphs[3].Text == "Text 3", "Paragraph text match");

                Assert.True(document.Lists[0].ListItems[0].Text == "Text 1", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[1].Text == "Text 2", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[2].Text == "Text 3", "Paragraph text match");

                Assert.True(document.Lists.Count == 1, "List count matches");

                paragraph = document.AddParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(WordListStyle.ArticleSections);
                wordList2.AddItem("Temp 2");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);

                Assert.True(document.Lists.Count == 2, "List count matches");

                paragraph = document.AddParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(WordListStyle.BulletedChars);
                wordList3.AddItem("Text 3");
                wordList3.AddItem("Text 2", 1);
                wordList3.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(WordListStyle.Headings111);
                wordList4.AddItem("Text 4");
                wordList4.AddItem("Text 2", 1);
                wordList4.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList5 = document.AddList(WordListStyle.Headings111Shifted);
                wordList5.AddItem("Text 5");
                wordList5.AddItem("Text 2", 1);
                wordList5.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 6th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList6 = document.AddList(WordListStyle.Chapters);
                wordList6.AddItem("Text 6");
                wordList6.AddItem("Text 2");
                wordList6.AddItem("Text 3");

                paragraph = document.AddParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList7 = document.AddList(WordListStyle.HeadingIA1);
                wordList7.AddItem("Text 7");
                wordList7.AddItem("Text 2", 1);
                wordList7.AddItem("Text 3", 2);

                Assert.True(document.Lists.Count == 7, "List count matches");

                Assert.True(document.Paragraphs.Count == 28, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 28, "Number of paragraphs on 1st section is wrong.");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx"))) {
                Assert.True(document.Paragraphs[0].IsListItem == false, "Paragraph is empty");
                Assert.True(document.Paragraphs[1].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[2].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[3].IsListItem == true, "Paragraph is list item 2");
                Assert.True(document.Paragraphs[3].Text == "Text 3", "Paragraph text match");

                Assert.True(document.Lists[0].ListItems[0].Text == "Text 1", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[1].Text == "Text 2", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[2].Text == "Text 3", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[2].ListItemLevel == 2, "Level doesn't match");

                document.Lists[0].ListItems[2].ListItemLevel = 1;
                document.Lists[0].ListItems[2].Text = "Text 4";

                Assert.True(document.Lists[0].ListItems[2].ListItemLevel == 1, "Level doesn't match");

                var paragraph = document.AddParagraph("This is 9th list").SetColor(Color.MediumAquamarine).SetUnderline(UnderlineValues.Double);

                WordList wordList8 = document.AddList(WordListStyle.Bulleted);
                wordList8.AddItem("Text 9");
                wordList8.AddItem("Text 9.1", 1);
                wordList8.AddItem("Text 9.2", 2);
                wordList8.AddItem("Text 9.3", 2);
                wordList8.AddItem("Text 9.4", 0);
                wordList8.AddItem("Text 9.5", 0);
                wordList8.AddItem("Text 9.6", 1);

                paragraph = document.AddParagraph("This is 10th list").SetColor(Color.ForestGreen).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(WordListStyle.Headings111);
                wordList2.AddItem("Temp 10");
                wordList2.AddItem("Text 10.1", 1);

                paragraph = document.AddParagraph("Paragraph in the middle of the list").SetColor(Color.Aquamarine); //.SetUnderline(UnderlineValues.Double);

                wordList2.AddItem("Text 10.2", 2);
                wordList2.AddItem("Text 10.3", 2);

                paragraph = document.AddParagraph("This is 10th list").SetColor(Color.ForestGreen).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(WordListStyle.Headings111);
                wordList3.AddItem("Temp 11");
                wordList3.AddItem("Text 11.1", 1);

                Assert.True(document.Lists.Count == 10, "List count matches");

                Assert.True(document.Paragraphs.Count == 45, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 45, "Number of paragraphs on 1st section is wrong.");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx"))) {
                Assert.True(document.Paragraphs[0].IsListItem == false, "Paragraph is empty");
                Assert.True(document.Paragraphs[1].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[2].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[3].IsListItem == true, "Paragraph is list item 2");


                Assert.True(document.Lists[0].ListItems[0].Text == "Text 1", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[1].Text == "Text 2", "Paragraph text match");
                // should work after reloading after last save
                Assert.True(document.Lists[0].ListItems[2].ListItemLevel == 1, "Level doesn't match");
                Assert.True(document.Lists[0].ListItems[2].Text == "Text 4", "Paragraph text match");
                Assert.True(document.Paragraphs[3].Text == "Text 4", "Paragraph text match");
                // We continue with the rest

                Assert.True(document.Lists.Count == 10, "List count matches");

                Assert.True(document.Paragraphs.Count == 45, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[0].Paragraphs.Count == 45, "Number of paragraphs on 1st section is wrong.");
                document.Save();
            }
        }
    }
}