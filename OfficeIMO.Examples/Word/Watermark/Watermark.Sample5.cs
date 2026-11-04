using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        /// <summary>
        /// Demonstrates applying watermarks with OfficeIMO colors and hex values across multiple sections.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_Sample5(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with mixed watermark color inputs");
            string filePath = Path.Combine(folderPath, "Watermark Multiple Colors.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // Section 0 - OfficeColor.Red
                document.Sections[0].PageSetup.Margins.Type = WordMargin.Normal;
                var watermark = document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Red");
                watermark.Color = Color.Red;
                Console.WriteLine($"Section 0 hex: {watermark.Color}");

                // Section 1 - OfficeColor.Green
                document.AddSection();
                document.Sections[1].PageSetup.Margins.Type = WordMargin.Normal;
                watermark = document.Sections[1].AddWatermark(WordWatermarkStyle.Text, "Green");
                watermark.Color = Color.Green;
                Console.WriteLine($"Section 1 hex: {watermark.Color}");

                // Section 2 - OfficeColor.Blue
                document.AddSection();
                document.Sections[2].PageSetup.Margins.Type = WordMargin.Normal;
                watermark = document.Sections[2].AddWatermark(WordWatermarkStyle.Text, "Blue");
                watermark.Color = Color.Blue;
                Console.WriteLine($"Section 2 hex: {watermark.Color}");

                // Section 3 - Hex without '#'
                document.AddSection();
                document.Sections[3].PageSetup.Margins.Type = WordMargin.Moderate;
                watermark = document.Sections[3].AddWatermark(WordWatermarkStyle.Text, "Magenta");
                watermark.Color = OfficeColor.Parse("ff00ff");

                // Section 4 - Hex with '#'
                document.AddSection();
                document.Sections[4].PageSetup.Margins.Type = WordMargin.Moderate;
                watermark = document.Sections[4].AddWatermark(WordWatermarkStyle.Text, "Cyan");
                watermark.Color = OfficeColor.Parse("#00ffff");

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
