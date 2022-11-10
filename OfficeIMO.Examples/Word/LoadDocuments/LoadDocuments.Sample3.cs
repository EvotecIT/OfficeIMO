using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class LoadDocuments {
        public static void LoadWordDocument_Sample3(bool openWord) {
            Console.WriteLine("[*] Load external Word Document - Sample 3");
            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "sample3.docx"), false)) {
                Console.WriteLine("Sections count: " + document.Sections.Count);
                Console.WriteLine("Tables count: " + document.Tables.Count);
                Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);
                document.Save(openWord);
            }
        }
    }
}
