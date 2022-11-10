using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {

        public static void EasyExample(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with HyperLink");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with HyperLink.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var para = document.AddHyperLink("Test", new Uri("https://evotec.xyz"));

                document.Save(openWord);
                Console.WriteLine("IsValid: " + document.DocumentIsValid);
                Console.WriteLine("Paragraph count: " + document.Paragraphs.Count);
            }
        }
    }
}
