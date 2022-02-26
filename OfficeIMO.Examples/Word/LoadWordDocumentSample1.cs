using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal class LoadWordDocumentSample1 {
        public static void LoadWordDocument_Sample1(bool openWord) {
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "sample1.docx"), true)) {


                Console.WriteLine("Sections count: " + document.Sections.Count);
                Console.WriteLine("Tables count: " + document.Tables.Count);
                Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);
                document.Save(openWord);
            }
        }
    }
}
