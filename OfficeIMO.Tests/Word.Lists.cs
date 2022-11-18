using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_CreatingWordDocumentWithLists() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx");
        using (var document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("First List");
            paragraph.ParagraphAlignment = JustificationValues.Center;

            Assert.False(document.Paragraphs[0].IsEmpty, "Paragraph is empty");

            Assert.Empty(document.Lists);

            var wordList1 = document.AddList(WordListStyle.Heading1ai);
            wordList1.AddItem("Text 1");
            wordList1.AddItem("Text 2", 1);
            wordList1.AddItem("Text 3", 2);

            Assert.False(document.Paragraphs[0].IsListItem, "Paragraph is empty");
            Assert.True(document.Paragraphs[1].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[2].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[3].IsListItem, "Paragraph is list item 2");
            Assert.Equal("Text 3", document.Paragraphs[3].Text);

            Assert.Equal("Text 1", document.Lists[0].ListItems[0].Text);
            Assert.Equal("Text 2", document.Lists[0].ListItems[1].Text);
            Assert.Equal("Text 3", document.Lists[0].ListItems[2].Text);

            Assert.Single(document.Lists);

            paragraph = document
                .AddParagraph("This is second list")
                .SetColor(Color.OrangeRed)
                .SetUnderline(UnderlineValues.Double);

            var wordList2 = document.AddList(WordListStyle.ArticleSections);
            wordList2.AddItem("Temp 2");
            wordList2.AddItem("Text 2", 1);
            wordList2.AddItem("Text 3", 2);

            Assert.Equal(2, document.Lists.Count);

            paragraph = document
                .AddParagraph("This is third list")
                .SetColor(Color.Blue)
                .SetUnderline(UnderlineValues.Double);

            var wordList3 = document.AddList(WordListStyle.BulletedChars);
            wordList3.AddItem("Text 3");
            wordList3.AddItem("Text 2", 1);
            wordList3.AddItem("Text 3", 2);

            paragraph = document
                .AddParagraph("This is fourth list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            var wordList4 = document.AddList(WordListStyle.Headings111);
            wordList4.AddItem("Text 4");
            wordList4.AddItem("Text 2", 1);
            wordList4.AddItem("Text 3", 2);

            paragraph = document
                .AddParagraph("This is five list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            WordList wordList5 = document.AddList(WordListStyle.Headings111Shifted);
            wordList5.AddItem("Text 5");
            wordList5.AddItem("Text 2", 1);
            wordList5.AddItem("Text 3", 2);

            paragraph = document
                .AddParagraph("This is 6th list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            var wordList6 = document.AddList(WordListStyle.Chapters);
            wordList6.AddItem("Text 6");
            wordList6.AddItem("Text 2");
            wordList6.AddItem("Text 3");

            paragraph = document
                .AddParagraph("This is 7th list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            var wordList7 = document.AddList(WordListStyle.HeadingIA1);
            wordList7.AddItem("Text 7");
            wordList7.AddItem("Text 2", 1);
            wordList7.AddItem("Text 3", 2);

            Assert.Equal(7, document.Lists.Count);
            Assert.Equal(28, document.Paragraphs.Count);

            var section = Assert.Single(document.Sections);
            Assert.Equal(28, section.Paragraphs.Count);

            document.Save(false);
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx"))) {
            Assert.False(document.Paragraphs[0].IsListItem, "Paragraph is empty");
            Assert.True(document.Paragraphs[1].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[2].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[3].IsListItem, "Paragraph is list item 2");
            Assert.Equal("Text 3", document.Paragraphs[3].Text);
            Assert.Equal("Text 1", document.Lists[0].ListItems[0].Text);
            Assert.Equal("Text 2", document.Lists[0].ListItems[1].Text);
            Assert.Equal("Text 3", document.Lists[0].ListItems[2].Text);
            Assert.Equal(2, document.Lists[0].ListItems[2].ListItemLevel);

            document.Lists[0].ListItems[2].ListItemLevel = 1;
            document.Lists[0].ListItems[2].Text = "Text 4";

            Assert.Equal(1, document.Lists[0].ListItems[2].ListItemLevel);

            var paragraph = document
                .AddParagraph("This is 9th list")
                .SetColor(Color.MediumAquamarine)
                .SetUnderline(UnderlineValues.Double);

            var wordList8 = document.AddList(WordListStyle.Bulleted);
            wordList8.AddItem("Text 9");
            wordList8.AddItem("Text 9.1", 1);
            wordList8.AddItem("Text 9.2", 2);
            wordList8.AddItem("Text 9.3", 2);
            wordList8.AddItem("Text 9.4", 0);
            wordList8.AddItem("Text 9.5", 0);
            wordList8.AddItem("Text 9.6", 1);

            paragraph = document
                .AddParagraph("This is 10th list")
                .SetColor(Color.ForestGreen)
                .SetUnderline(UnderlineValues.Double);

            var wordList2 = document.AddList(WordListStyle.Headings111);
            wordList2.AddItem("Temp 10");
            wordList2.AddItem("Text 10.1", 1);

            paragraph = document
                .AddParagraph("Paragraph in the middle of the list")
                .SetColor(Color.Aquamarine); //.SetUnderline(UnderlineValues.Double);

            wordList2.AddItem("Text 10.2", 2);
            wordList2.AddItem("Text 10.3", 2);

            paragraph = document
                .AddParagraph("This is 10th list")
                .SetColor(Color.ForestGreen).SetUnderline(UnderlineValues.Double);

            var wordList3 = document.AddList(WordListStyle.Headings111);
            wordList3.AddItem("Temp 11");
            wordList3.AddItem("Text 11.1", 1);

            Assert.Equal(10, document.Lists.Count);
            Assert.Equal(45, document.Paragraphs.Count);

            var section = Assert.Single(document.Sections);
            Assert.Equal(45, section.Paragraphs.Count);

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists.docx"))) {
            Assert.False(document.Paragraphs[0].IsListItem, "Paragraph is empty");
            Assert.True(document.Paragraphs[1].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[2].IsListItem, "Paragraph is list item 1");
            Assert.True(document.Paragraphs[3].IsListItem, "Paragraph is list item 2");

            Assert.Equal("Text 1", document.Lists[0].ListItems[0].Text);
            Assert.Equal("Text 2", document.Lists[0].ListItems[1].Text);
            // should work after reloading after last save
            Assert.Equal(1, document.Lists[0].ListItems[2].ListItemLevel);
            Assert.Equal("Text 4", document.Lists[0].ListItems[2].Text);
            Assert.Equal("Text 4", document.Paragraphs[3].Text);
            // We continue with the rest
            Assert.Equal(10, document.Lists.Count);
            Assert.Equal(45, document.Paragraphs.Count);

            var section = Assert.Single(document.Sections);
            Assert.Equal(45, section.Paragraphs.Count);

            document.Save();
        }
    }

    [Fact]
    public void Test_CreatingWordDocumentWithLists2() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists2.docx");
        using (var document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 4");
            paragraph.ParagraphAlignment = JustificationValues.Center;

            var wordList = document.AddList(WordListStyle.Headings111);
            wordList.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList.AddItem("Text 2.1", 1).SetColor(Color.Brown);
            wordList.AddItem("Text 2.2", 1).SetColor(Color.Brown);
            wordList.AddItem("Text 2.3", 1).SetColor(Color.Brown);
            wordList.AddItem("Text 2.3.4", 2).SetColor(Color.Brown);
            // here we set another list element but we also change it using standard paragraph change
            paragraph = wordList.AddItem("Text 3");
            paragraph.Bold = true;
            paragraph.SetItalic();

            Assert.Single(document.Lists);

            paragraph = document
                .AddParagraph("This is second list")
                .SetColor(Color.OrangeRed)
                .SetUnderline(UnderlineValues.Double);

            WordList wordList1 = document.AddList(WordListStyle.HeadingIA1);
            wordList1.AddItem("Temp 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList1.AddItem("Temp 2.1", 1).SetColor(Color.Brown);
            wordList1.AddItem("Temp 2.2", 1).SetColor(Color.Brown);
            wordList1.AddItem("Temp 2.3", 1).SetColor(Color.Brown);
            wordList1.AddItem("Temp 2.3.4", 2).SetColor(Color.Brown).Remove();
            wordList1.ListItems[1].Remove();
            paragraph = wordList1.AddItem("Temp 3");

            Assert.Equal(2, document.Lists.Count);
            Assert.Equal(2, document.Sections[0].Lists.Count);

            paragraph = document
                .AddParagraph("This is third list")
                .SetColor(Color.Blue)
                .SetUnderline(UnderlineValues.Double);

            var wordList2 = document.AddList(WordListStyle.BulletedChars);
            wordList2.AddItem("Oops 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList2.AddItem("Oops 2.1", 1).SetColor(Color.Brown);
            wordList2.AddItem("Oops 2.2", 1).SetColor(Color.Brown);
            wordList2.AddItem("Oops 2.3", 1).SetColor(Color.Brown);
            wordList2.AddItem("Oops 2.3.4", 2).SetColor(Color.Brown);

            Assert.Equal(3, document.Lists.Count);
            Assert.Equal(3, document.Sections[0].Lists.Count);

            paragraph = document
                .AddParagraph("This is fourth list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            var wordList3 = document.AddList(WordListStyle.Heading1ai);
            wordList3.AddItem("4th 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList3.AddItem("4th 2.1", 1).SetColor(Color.Brown);
            wordList3.AddItem("4th 2.2", 1).SetColor(Color.Brown);
            wordList3.AddItem("4th 2.3", 1).SetColor(Color.Brown);
            wordList3.AddItem("4th 2.3.4", 2).SetColor(Color.Brown);

            Assert.Equal(4, document.Lists.Count);
            Assert.Equal(4, document.Sections[0].Lists.Count);

            paragraph = document
                .AddParagraph("This is five list")
                .SetColor(Color.DeepPink)
                .SetUnderline(UnderlineValues.Double);

            var wordList4 = document.AddList(WordListStyle.Headings111Shifted);
            wordList4.AddItem("5th 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList4.AddItem("5th 2.1", 1).SetColor(Color.Brown);
            wordList4.AddItem("5th 2.2", 1).SetColor(Color.Brown);
            wordList4.AddItem("5th 2.3", 1).SetColor(Color.Brown);
            wordList4.AddItem("5th 2.3.4", 2).SetColor(Color.Brown);

            Assert.Equal(5, document.Lists.Count);
            Assert.Equal(6, document.Lists[0].ListItems.Count);
            Assert.Equal(4, document.Lists[1].ListItems.Count);
            Assert.Equal(5, document.Lists[2].ListItems.Count);
            Assert.Equal(5, document.Lists[3].ListItems.Count);
            Assert.Equal(5, document.Lists[4].ListItems.Count);

            document.Lists[3].ListItems[2].Text = "Overwrite Text 2.2";
            document.Lists[4].ListItems[2].Text = "Overwrite Text 2.12";

            Assert.Equal(5, document.Lists.Count);

            document.Lists[3].AddItem("Added 2.3.5", 3).SetColor(Color.DimGrey);
            document.Lists[2].AddItem("Added 2.3.5", 3).SetColor(Color.DimGrey);

            Assert.Equal(5, document.Lists.Count);

            Assert.Equal("Text 1", document.Lists[0].ListItems[0].Text);
            Assert.Equal("Text 2.1", document.Lists[0].ListItems[1].Text);
            Assert.Equal("Temp 2.2", document.Lists[1].ListItems[1].Text);
            Assert.Equal("Temp 2.3", document.Lists[1].ListItems[2].Text);

            Assert.Equal("Oops 2.1", document.Lists[2].ListItems[1].Text);
            Assert.Equal("Oops 2.2", document.Lists[2].ListItems[2].Text);

            Assert.Equal("Overwrite Text 2.2", document.Lists[3].ListItems[2].Text);
            Assert.Equal("Overwrite Text 2.12", document.Lists[4].ListItems[2].Text);

            Assert.Equal(6, document.Lists[2].ListItems.Count);
            Assert.Equal(6, document.Lists[3].ListItems.Count);

            var section = document.AddSection();
            section.PageSettings.Orientation = PageOrientationValues.Landscape;

            Assert.Equal(PageOrientationValues.Portrait, document.Sections[0].PageSettings.Orientation);
            Assert.Equal(PageOrientationValues.Landscape, document.Sections[1].PageSettings.Orientation);

            Assert.Equal(5, document.Lists.Count);
            Assert.Equal(5, document.Sections[0].Lists.Count);

            Assert.Empty(document.Sections[1].Lists);

            var wordList5 = document.AddList(WordListStyle.Headings111);
            wordList5.AddItem("Section 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList5.AddItem("Section 2.1", 1).SetColor(Color.Brown);
            wordList5.AddItem("Section 2.2", 1).SetColor(Color.Brown);

            Assert.Equal(6, document.Lists.Count);
            Assert.Equal(5, document.Sections[0].Lists.Count);
            Assert.Single(document.Sections[1].Lists);

            document.Save(false);
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists2.docx"))) {
            Assert.Equal(6, document.Lists.Count);
            Assert.Equal(5, document.Sections[0].Lists.Count);
            Assert.Single(document.Sections[1].Lists);

            var wordList6 = document.AddList(WordListStyle.Headings111);
            wordList6.AddItem("Section 1").SetCapsStyle(CapsStyle.SmallCaps);
            wordList6.AddItem("Section 2.1", 1).SetColor(Color.Brown);
            wordList6.AddItem("Section 2.2", 1).SetColor(Color.Brown);

            Assert.Equal(7, document.Lists.Count);
            Assert.Equal(5, document.Sections[0].Lists.Count);
            Assert.Equal(2, document.Sections[1].Lists.Count);

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLists2.docx"))) {
            Assert.Equal(7, document.Lists.Count);
            Assert.Equal(5, document.Sections[0].Lists.Count);
            Assert.Equal(2, document.Sections[1].Lists.Count);
            document.Save();
        }
    }

    [Fact]
    public void Test_SavingWordDocumentWithListsToStream() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithListsToStream.docx");
        var wordListStyles = (WordListStyle[]) Enum.GetValues(typeof(WordListStyle));
        using (var document = WordDocument.Create()) {
            foreach (var listStyle in wordListStyles) {
                var paragraph = document.AddParagraph(listStyle.ToString());
                paragraph.SetColor(Color.Red).SetBold();
                paragraph.ParagraphAlignment = JustificationValues.Center;

                var wordList1 = document.AddList(listStyle);
                wordList1.AddItem("Text 1");
            }

            using var outputStream = new MemoryStream();
            document.Save(outputStream);
            File.WriteAllBytes(filePath, outputStream.ToArray());
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Equal(wordListStyles.Length, document.Lists.Count);
            var abstractNums = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!
                .Numbering.ChildElements.OfType<AbstractNum>().ToArray();
            for (var idx = 0; idx < abstractNums.Length; idx++) {
                var style = WordListStyles.GetStyle(wordListStyles[idx]);
                Assert.Equal(style, abstractNums[idx]);
            }
        }
    }
}
