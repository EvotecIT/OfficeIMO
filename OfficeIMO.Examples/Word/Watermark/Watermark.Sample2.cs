using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_Sample2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with watermark");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with watermark.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                // we skip headers/footers which will be created for us

                document.Sections[0].SetMargins(WordMargin.Normal);

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Confidential");

                document.AddSection();

                document.Sections[1].SetMargins(WordMargin.Moderate);

                Console.WriteLine("----");

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                Console.WriteLine(document.Sections[1].Margins.Left.Value);
                Console.WriteLine(document.Sections[1].Margins.Right.Value);

                document.Settings.SetBackgroundColor(Color.Azure);

                document.Save(openWord);
            }
        }
    }
}
