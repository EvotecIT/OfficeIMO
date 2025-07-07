using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        /// <summary>
        /// Removes all watermarks from a document.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_Remove(string folderPath, bool openWord) {
            Console.WriteLine("[*] Removing watermarks using RemoveWatermark");
            string filePath = System.IO.Path.Combine(folderPath, "Watermark Remove.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Default");
                document.Sections[0].Header.First.AddWatermark(WordWatermarkStyle.Text, "First");
                document.Sections[0].Header.Even.AddWatermark(WordWatermarkStyle.Text, "Even");

                Console.WriteLine("Watermarks before: " + document.Watermarks.Count);
                foreach (var watermark in document.Watermarks.ToList()) {
                    watermark.Remove();
                }
                Console.WriteLine("Watermarks after: " + document.Watermarks.Count);

                document.Save(openWord);
            }
        }
    }
}
