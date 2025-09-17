using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CoverPages {
        public static void Example_AddingCoverPage(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Cover Page");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Cover Page.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].PageSettings.PageSize = WordPageSize.A4;

                document.PageSettings.PageSize = WordPageSize.A4;

                var pageSettings = document.PageSettings ?? throw new InvalidOperationException("Page settings are not initialized.");

                Console.WriteLine(pageSettings.Height?.ToString() ?? "Height not set");
                Console.WriteLine(pageSettings.Width?.ToString() ?? "Width not set");
                Console.WriteLine(pageSettings.Code?.ToString() ?? "Code not set");
                Console.WriteLine(pageSettings.PageSize?.ToString() ?? "Page size not set");

                document.BuiltinDocumentProperties.Title = "Cover Page Templates";
                document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";

                document.ApplicationProperties.Company = "Evotec Services";

                document.Settings.UpdateFieldsOnOpen = true;

                document.AddCoverPage(CoverPageTemplate.IonDark);

                document.AddTableOfContent(TableOfContentStyle.Template1);

                document.AddPageBreak();

                var wordListToc = document.AddTableOfContentList(WordListStyle.Numbered);

                wordListToc.AddItem("Prepare document");

                document.AddParagraph("This is my test 1");

                wordListToc.AddItem("Make it shine");

                document.AddParagraph("This is my test 2");

                document.AddPageBreak();

                wordListToc.AddItem("More on the next page");

                document.Save(openWord);
            }
        }
    }
}
