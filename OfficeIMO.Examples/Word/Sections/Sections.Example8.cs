using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {
        internal static void Example_CloneSection(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning sections");
            string filePath = Path.Combine(folderPath, "Document with Cloned Section.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section0 - Paragraph0");
                var section1 = document.AddSection();
                section1.AddParagraph("Section1 - Paragraph0");
                document.AddParagraph("Section2 - Paragraph0");
                Console.WriteLine("+ Sections: " + document.Sections.Count);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var clone = document.Sections[0].CloneSection();
                Console.WriteLine("+ Sections: " + document.Sections.Count);
                Console.WriteLine("Section 0 P0: " + document.Sections[0].Paragraphs[0].Text);
                Console.WriteLine("Section 1 P0: " + document.Sections[1].Paragraphs[0].Text);
                Console.WriteLine("Section 2 P0: " + document.Sections[2].Paragraphs[0].Text);
                Console.WriteLine("+ Sections: " + document.Sections.Count);
                document.Save(openWord);
            }
        }
    }
}

