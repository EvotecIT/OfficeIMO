using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_Sample1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Watermark 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark 2.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                document.Sections[0].SetMargins(WordMargin.Normal);

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                Console.WriteLine(document.Sections[0].Margins.Type);

                document.Sections[0].Margins.Type = WordMargin.Wide;


                Console.WriteLine(document.Sections[0].Margins.Type);

                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Margins.Type = WordMargin.Narrow;

                Console.WriteLine("----");
                document.Sections[1].AddWatermark(WordWatermarkStyle.Text, "Draft");

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                Console.WriteLine(document.Sections[1].Margins.Left.Value);
                Console.WriteLine(document.Sections[1].Margins.Right.Value);

                Console.WriteLine(document.Sections[1].Margins.Type);


                document.Settings.SetBackgroundColor(Color.Azure);

                document.Save(openWord);
            }
        }
    }
}
