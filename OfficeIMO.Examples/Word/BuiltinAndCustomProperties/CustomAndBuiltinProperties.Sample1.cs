using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class CustomAndBuiltinProperties {

        public static void Example_LoadDocumentWithProperties(bool openWord = false) {
            Console.WriteLine("[*] Loading standard document to check properties");

            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string filePath = Path.Combine(folderPath, "DocumentWithBuiltinAndCustomProperties.docx");

            using (WordDocument document = WordDocument.Load(filePath, true)) {
                Console.WriteLine("+ Document Path: " + document.FilePath);
                Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
                Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);

                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                Console.WriteLine(document.ApplicationProperties.ApplicationVersion);

                document.Open(openWord);
            }
        }
    }
}
