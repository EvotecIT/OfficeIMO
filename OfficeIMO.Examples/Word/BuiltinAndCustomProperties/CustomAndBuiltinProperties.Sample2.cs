using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CustomAndBuiltinProperties {
        /// <summary>
        /// Loads a document and prints basic property information.
        /// </summary>
        /// <param name="openWord">Whether to open Word after loading the document.</param>
        public static void Example_Load(bool openWord = false) {
            Console.WriteLine("[*] Loading basic document");

            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string filePath = Path.Combine(folderPath, "DocumentWithSection.docx");
            //filePath = @"C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Documents\DocumentWithSection.docx";
            //filePath = @"C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Documents\EmptyDocumentWithSection.docx";

            using (WordDocument document = WordDocument.Load(filePath, true)) {
                Console.WriteLine("+ Document Path: " + document.FilePath);
                Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
                Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);

                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                document.Open(openWord);
            }
        }

    }
}
