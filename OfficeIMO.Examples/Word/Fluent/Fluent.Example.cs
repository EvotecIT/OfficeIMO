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
                    document.Properties.Title = "Fluent Document";
                    document.Properties.Creator = "Fluent Author";
                    document.Properties.Subject = "Fluent Subject";
                    document.Properties.Keywords = "fluent, api";
                    document.Properties.Description = "Created via fluent API";
                    document.CustomProperties["Reviewed"] = new WordCustomProperty(true);
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