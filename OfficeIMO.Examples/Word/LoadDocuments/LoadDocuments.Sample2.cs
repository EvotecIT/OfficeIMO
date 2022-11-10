using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class LoadDocuments {
        public static void LoadWordDocument_Sample2(bool openWord) {
            Console.WriteLine("[*] Load external Word Document - Sample 2");

            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "sample2.docx"), false)) {


                Console.WriteLine("Sections count: " + document.Sections.Count);
                Console.WriteLine("Tables count: " + document.Tables.Count);
                Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);
                document.Save(openWord);
            }
        }
    }
}
