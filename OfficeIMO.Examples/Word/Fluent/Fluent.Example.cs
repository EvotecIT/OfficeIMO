using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentDocument(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent API");
            string filePath = Path.Combine(folderPath, "FluentDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    document.BuiltinDocumentProperties.Title = "Fluent Document";
                    document.BuiltinDocumentProperties.Creator = "Fluent Author";
                    document.BuiltinDocumentProperties.Subject = "Fluent Subject";
                    document.BuiltinDocumentProperties.Keywords = "fluent, api";
                    document.BuiltinDocumentProperties.Description = "Created via fluent API";
                    document.CustomDocumentProperties["Reviewed"] = new WordCustomProperty(true);
                    var section1 = document.Sections.Last();
                    section1 = document.AddSection();
                    var paragraph2 = document.AddParagraph();
                    paragraph2 = paragraph2.AddText("Hello from fluent API");
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}