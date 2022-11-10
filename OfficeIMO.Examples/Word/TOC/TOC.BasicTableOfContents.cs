using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class TOC {

        internal static void Example_BasicTOC1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with TOC - 1");
            string filePath = System.IO.Path.Combine(folderPath, "Document with TOC1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                // Standard way to open document and be asked about Updating Fields including TOC
                document.Settings.UpdateFieldsOnOpen = true;

                WordTableOfContent wordTableContent = document.AddTableOfContent(TableOfContentStyle.Template1);
                wordTableContent.Text = "This is Table of Contents";
                wordTableContent.TextNoContent = "Ooopsi, no content";

                document.AddPageBreak();

                var paragraph = document.AddParagraph("Test");
                paragraph.Style = WordParagraphStyles.Heading1;

                Console.WriteLine(wordTableContent.Text);
                Console.WriteLine(wordTableContent.TextNoContent);

                //// i am not sure if this is even working properly, seems so, but seems bad idea
                //wordTableContent.Update();

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(openWord);
            }
        }

    }
}
