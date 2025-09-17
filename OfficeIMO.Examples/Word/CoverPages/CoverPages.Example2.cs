using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CoverPages {
        public static void Example_AddingCoverPage2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Cover Page");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Cover Page.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {

                document.BuiltinDocumentProperties.Title = "Cover Page Templates";
                document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";

                document.Settings.UpdateFieldsOnOpen = true;

                document.AddCoverPage(CoverPageTemplate.Austin);

                var tableOfContent = document.AddTableOfContent();

                document.AddPageBreak();

                var wordListToc = document.AddTableOfContentList(WordListStyle.Numbered);

                wordListToc.AddItem("Prepare document");

                document.AddParagraph("This is my test 1");

                wordListToc.AddItem("Make it shine");

                document.AddParagraph("This is my test 2");

                document.AddPageBreak();

                wordListToc.AddItem("More on the next page");

                tableOfContent.Update();

                document.Save(openWord);
            }
        }
    }
}
