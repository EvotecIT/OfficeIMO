using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithLists() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.InsertParagraph("First List");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                Assert.True(document.Paragraphs[0].IsEmpty == false, "Paragraph is empty");

                Assert.True(document.Lists.Count == 0, "List count matches");

                WordList wordList1 = document.AddList(ListStyles.Heading1ai);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                Assert.True(document.Paragraphs[0].IsListItem == false, "Paragraph is empty");
                Assert.True(document.Paragraphs[1].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[2].IsListItem == true, "Paragraph is list item 1");
                Assert.True(document.Paragraphs[3].IsListItem == true, "Paragraph is list item 2");
                Assert.True(document.Paragraphs[3].Text == "Text 3" ,"Paragraph text match");

                Assert.True(document.Lists[0].ListItems[0].Text == "Text 1", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[1].Text == "Text 2", "Paragraph text match");
                Assert.True(document.Lists[0].ListItems[2].Text == "Text 3", "Paragraph text match");

                Assert.True(document.Lists.Count == 1, "List count matches");

                paragraph = document.InsertParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(ListStyles.ArticleSections);
                wordList2.AddItem("Temp 2");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);

                Assert.True(document.Lists.Count == 2, "List count matches");

                paragraph = document.InsertParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(ListStyles.BulletedChars);
                wordList3.AddItem("Text 3");
                wordList3.AddItem("Text 2", 1);
                wordList3.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(ListStyles.Headings111);
                wordList4.AddItem("Text 4");
                wordList4.AddItem("Text 2", 1);
                wordList4.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList5 = document.AddList(ListStyles.Headings111Shifted);
                wordList5.AddItem("Text 5");
                wordList5.AddItem("Text 2", 1);
                wordList5.AddItem("Text 3", 2);

                paragraph = document.InsertParagraph("This is 6th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList6 = document.AddList(ListStyles.Chapters);
                wordList6.AddItem("Text 6");
                wordList6.AddItem("Text 2");
                wordList6.AddItem("Text 3");

                paragraph = document.InsertParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList7 = document.AddList(ListStyles.HeadingIA1);
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

            }
        } }
}
