using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        /// <summary>
        /// Demonstrates applying watermarks with SixLabors colors and hex values across multiple sections.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_Sample5(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with mixed watermark color inputs");
            string filePath = Path.Combine(folderPath, "Watermark Multiple Colors.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // Section 0 - SixLabors Color.Red
                document.Sections[0].SetMargins(WordMargin.Normal);
                var watermark = document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Red");
                watermark.Color = Color.Red;
                Console.WriteLine($"Section 0 hex: {watermark.ColorHex}");

                // Section 1 - SixLabors Color.Green
                document.AddSection();
                document.Sections[1].SetMargins(WordMargin.Normal);
                watermark = document.Sections[1].AddWatermark(WordWatermarkStyle.Text, "Green");
                watermark.Color = Color.Green;
                Console.WriteLine($"Section 1 hex: {watermark.ColorHex}");

                // Section 2 - SixLabors Color.Blue
                document.AddSection();
                document.Sections[2].SetMargins(WordMargin.Normal);
                watermark = document.Sections[2].AddWatermark(WordWatermarkStyle.Text, "Blue");
                watermark.Color = Color.Blue;
                Console.WriteLine($"Section 2 hex: {watermark.ColorHex}");

                // Section 3 - Hex without '#'
                document.AddSection();
                document.Sections[3].SetMargins(WordMargin.Moderate);
                watermark = document.Sections[3].AddWatermark(WordWatermarkStyle.Text, "Magenta");
                watermark.ColorHex = "ff00ff";

                // Section 4 - Hex with '#'
                document.AddSection();
                document.Sections[4].SetMargins(WordMargin.Moderate);
                watermark = document.Sections[4].AddWatermark(WordWatermarkStyle.Text, "Cyan");
                watermark.ColorHex = "#00ffff";

                document.Save(openWord);
            }
        }
    }
}