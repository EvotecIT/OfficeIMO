using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        /// <summary>
        /// Creates a document with an image watermark.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_SampleImage1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Watermark Image 1");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark Image 2.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();

                var imagePathToAdd = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
                var section = document.Sections[0];
                var header = GetSectionHeaderOrThrow(section);
                var watermark = header.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd);

                //Console.WriteLine(watermark.Height);
                //Console.WriteLine(watermark.Width);

                //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);
                //Console.WriteLine("Images in document: " + document.Images.Count);
                //Console.WriteLine("Watermarks in header: " + document.Header!.Default.Watermarks.Count);
                //Console.WriteLine("Images in header: " + document.Header!.Default.Images.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);
                //Console.WriteLine("Images in document: " + document.Images.Count);
                //Console.WriteLine("Watermarks in header: " + document.Header!.Default.Watermarks.Count);
                //Console.WriteLine("Images in header: " + document.Header!.Default.Images.Count);
                document.Save(openWord);
            }
        }

    }
}
