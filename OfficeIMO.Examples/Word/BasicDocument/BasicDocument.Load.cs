using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicLoadHamlet(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading Hamlet Document");
            string filePath = System.IO.Path.Combine(templatesPath, "Hamlet.docx");

            using (WordDocument document = WordDocument.Load(filePath)) {
                // TODO: add reading/writing FootnoteProperties/EndnoteProperties 

                Console.WriteLine("----");
                Console.WriteLine(document.Sections.Count);
                Console.WriteLine("----");
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);

                Console.WriteLine(document.Sections[0].HyperLinks.Count);
                Console.WriteLine(document.HyperLinks.Count);
                Console.WriteLine(document.Fields.Count);
                document.Save(openWord);
            }
        }
    }
}
