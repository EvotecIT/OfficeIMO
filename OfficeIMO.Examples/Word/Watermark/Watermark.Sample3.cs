using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_Sample3(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with watermark");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with watermark and sections.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddParagraph("Section 0");
                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Confidential");

                document.AddPageBreak();
                document.AddPageBreak();

                var section = document.AddSection();
                section.AddWatermark(WordWatermarkStyle.Text, "Second Mark");

                document.AddParagraph("Section 1");

                document.AddPageBreak();
                document.AddPageBreak();

                var section1 = document.AddSection();

                document.AddParagraph("Section 2");

                document.Sections[2].AddWatermark(WordWatermarkStyle.Text, "New");

                document.AddPageBreak();
                document.AddPageBreak();

                Console.WriteLine("----");
                Console.WriteLine("Watermarks: " + document.Watermarks.Count);
                Console.WriteLine("Watermarks section 0: " + document.Sections[0].Watermarks.Count);
                Console.WriteLine("Watermarks section 1: " + document.Sections[1].Watermarks.Count);
                Console.WriteLine("Watermarks section 2: " + document.Sections[2].Watermarks.Count);

                Console.WriteLine("Paragraphs: " + document.Paragraphs.Count);

                Console.WriteLine("Removing last watermark");

                document.Sections[2].Watermarks[0].Remove();

                Console.WriteLine("Watermarks: " + document.Watermarks.Count);
                Console.WriteLine("Watermarks section 0: " + document.Sections[0].Watermarks.Count);
                Console.WriteLine("Watermarks section 1: " + document.Sections[1].Watermarks.Count);
                Console.WriteLine("Watermarks section 2: " + document.Sections[2].Watermarks.Count);
                Console.WriteLine("Paragraphs: " + document.Paragraphs.Count);

                document.Save(openWord);
            }
        }
    }
}
