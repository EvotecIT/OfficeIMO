using System;
using OfficeIMO.Examples.Utils;
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

                var firstSection = document.Sections[0];
                var headers = Guard.NotNull(firstSection.Header, "Headers should exist after calling AddHeadersAndFooters.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header should exist after calling AddHeadersAndFooters.");
                defaultHeader.AddWatermark(WordWatermarkStyle.Text, "Default");
                var firstHeader = Guard.NotNull(headers.First, "First header should exist after enabling different first page.");
                firstHeader.AddWatermark(WordWatermarkStyle.Text, "First");
                var evenHeader = Guard.NotNull(headers.Even, "Even header should exist after enabling different odd and even pages.");
                evenHeader.AddWatermark(WordWatermarkStyle.Text, "Even");

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
