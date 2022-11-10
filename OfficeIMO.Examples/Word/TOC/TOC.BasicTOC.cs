using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class TOC {

        internal static void Example_BasicTOC2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with TOC - 2");
            string filePath = System.IO.Path.Combine(folderPath, "Document with TOC2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                // Standard way to open document and be asked about Updating Fields including TOC
                document.Settings.UpdateFieldsOnOpen = true;

                WordTableOfContent wordTableContent = document.AddTableOfContent(TableOfContentStyle.Template1);
                wordTableContent.Text = "This is Table of Contents";
                wordTableContent.TextNoContent = "Ooopsi, no content";

                document.AddPageBreak();

                WordList wordList = document.AddList(WordListStyle.Headings111);
                wordList.AddItem("Text 1").Style = WordParagraphStyles.Heading1;

                document.AddPageBreak();

                wordList.AddItem("Text 2.1", 1).SetColor(Color.Brown).Style = WordParagraphStyles.Heading2;

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(openWord);
            }
        }

    }
}
