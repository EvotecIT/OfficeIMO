using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_SampleImage1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Watermark Image 1");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark Image 2.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();

                var imagePathToAdd = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
                var watermark = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd);

                //Console.WriteLine(watermark.Height);
                //Console.WriteLine(watermark.Width);

                //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);
                //Console.WriteLine("Images in document: " + document.Images.Count);
                //Console.WriteLine("Watermarks in header: " + document.Header.Default.Watermarks.Count);
                //Console.WriteLine("Images in header: " + document.Header.Default.Images.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);
                //Console.WriteLine("Images in document: " + document.Images.Count);
                //Console.WriteLine("Watermarks in header: " + document.Header.Default.Watermarks.Count);
                //Console.WriteLine("Images in header: " + document.Header.Default.Images.Count);
                document.Save(openWord);
            }
        }

    }
}
