using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class DocumentProperties
    {
        public static void Example_FluentDocumentProperties(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent document properties");
            string filePath = Path.Combine(folderPath, "Fluent Document Properties.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    document.BuiltinDocumentProperties.Title = "Title";
                    document.BuiltinDocumentProperties.Creator = "Author";
                    document.BuiltinDocumentProperties.Subject = "Subject";
                    document.BuiltinDocumentProperties.Keywords = "k1, k2";
                    document.BuiltinDocumentProperties.Description = "Some comments";
                    document.BuiltinDocumentProperties.Category = "Category";
                    document.ApplicationProperties.Company = "Evotec";
                    document.ApplicationProperties.Manager = "Manager1";
                    document.BuiltinDocumentProperties.LastModifiedBy = "John";
                    document.BuiltinDocumentProperties.Revision = "1.0";
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Test");
                }

                document.Save();
                if (openWord)
                    document.OpenInApplication();
            }
        }
    }
}