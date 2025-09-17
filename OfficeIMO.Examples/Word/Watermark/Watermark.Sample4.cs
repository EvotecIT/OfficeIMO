using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        /// <summary>
        /// Demonstrates how to apply a watermark using a hex color value.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_Sample4(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with Watermark hex color");
            string filePath = Path.Combine(folderPath, "Basic Document with Watermark Hex Color.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                var section0DefaultHeader = Guard.NotNull(document.Sections[0].Header?.Default, "Section 0 should expose a default header after adding headers and footers.");
                var watermark = section0DefaultHeader.AddWatermark(WordWatermarkStyle.Text, "HexColor");
                watermark.ColorHex = "00ff00";
                document.Save(openWord);
            }
        }
    }
}

