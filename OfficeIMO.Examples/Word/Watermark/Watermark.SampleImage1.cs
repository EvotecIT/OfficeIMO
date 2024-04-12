using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_SampleImage1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Watermark Image 1");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark Image 1.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();

                var imagePathToAdd = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
                var watermark = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd);
                watermark.Height = 100;
                watermark.Width = 100;



                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                document.Save(openWord);
            }
        }

    }
}
