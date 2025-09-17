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
        /// Demonstrates how to create a document with a basic watermark.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Whether to open the document after creation.</param>
        public static void Watermark_Sample1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Watermark 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Watermark 4.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                var section0DefaultHeader = Guard.NotNull(document.Sections[0].Header?.Default, "Section 0 should expose a default header after adding headers and footers.");
                section0DefaultHeader.AddParagraph("Section 0 - In header");
                document.Sections[0].SetMargins(WordMargin.Normal);

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                Console.WriteLine(document.Sections[0].Margins.Type);

                document.Sections[0].Margins.Type = WordMargin.Wide;

                Console.WriteLine(document.Sections[0].Margins.Type);

                Console.WriteLine("----");
                var watermark = section0DefaultHeader.AddWatermark(WordWatermarkStyle.Text, "Watermark");
                watermark.Color = Color.Red;

                // ColorHex normally returns hex colors, but for watermark it returns string as the underlying value is in string name, not hex
                Console.WriteLine(watermark.ColorHex);

                Console.WriteLine(watermark.Rotation);

                watermark.Rotation = 180;

                Console.WriteLine(watermark.Rotation);

                watermark.Stroked = true;

                Console.WriteLine(watermark.Height);
                Console.WriteLine(watermark.Width);

                // width and height in points (HTML wise)
                watermark.Height = 100.15;
                watermark.Width = 500.18;

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddSection();

                document.AddParagraph("Section 1");

                document.Sections[1].AddHeadersAndFooters();
                var section1DefaultHeader = Guard.NotNull(document.Sections[1].Header?.Default, "Section 1 should expose a default header after adding headers and footers.");
                section1DefaultHeader.AddParagraph("Section 1 - In header");
                document.Sections[1].Margins.Type = WordMargin.Narrow;
                Console.WriteLine("----");

                Console.WriteLine("Section 0 - Paragraphs Count: " + section0DefaultHeader.Paragraphs.Count);
                Console.WriteLine("Section 1 - Paragraphs Count: " + section1DefaultHeader.Paragraphs.Count);

                Console.WriteLine("----");
                document.Sections[1].AddParagraph("Test");
                section1DefaultHeader.AddWatermark(WordWatermarkStyle.Text, "Draft");

                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                Console.WriteLine(document.Sections[1].Margins.Left.Value);
                Console.WriteLine(document.Sections[1].Margins.Right.Value);

                Console.WriteLine(document.Sections[1].Margins.Type);


                document.Settings.SetBackgroundColor(Color.Azure);

                Console.WriteLine("----");

                var defaultHeader = Guard.NotNull(document.Header?.Default, "Document should expose a default header after adding headers and footers.");
                Console.WriteLine("Watermarks in default header: " + defaultHeader.Watermarks.Count);

                var defaultFooter = Guard.NotNull(document.Footer?.Default, "Document should expose a default footer after adding headers and footers.");
                Console.WriteLine("Watermarks in default footer: " + defaultFooter.Watermarks.Count);

                Console.WriteLine("Watermarks in section 0: " + document.Sections[0].Watermarks.Count);
                Console.WriteLine("Watermarks in section 0 (header): " + section0DefaultHeader.Watermarks.Count);
                Console.WriteLine("Paragraphs in section 0 (header): " + section0DefaultHeader.Paragraphs.Count);

                Console.WriteLine("Watermarks in section 1: " + document.Sections[1].Watermarks.Count);
                Console.WriteLine("Watermarks in section 1 (header): " + section1DefaultHeader.Watermarks.Count);
                Console.WriteLine("Paragraphs in section 1 (header): " + section1DefaultHeader.Paragraphs.Count);

                Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                //Console.WriteLine("----");
                //Console.WriteLine("Watermarks in default header: " + document.Header!.Default.Watermarks.Count);

                //Console.WriteLine("Watermarks in default footer: " + document.Footer!.Default.Watermarks.Count);

                //Console.WriteLine("Watermarks in section 0: " + document.Sections[0].Watermarks.Count);
                //Console.WriteLine("Watermarks in section 0 (header): " + document.Sections[0].Header!.Default.Watermarks.Count);
                //Console.WriteLine("Paragraphs in section 0 (header): " + document.Sections[0].Header!.Default.Paragraphs.Count);

                //Console.WriteLine("Watermarks in section 1: " + document.Sections[1].Watermarks.Count);

                //Console.WriteLine("Paragraphs in section 1 (header): " + document.Sections[1].Header!.Default.Paragraphs.Count);

                //Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);

                document.Save(openWord);
            }
        }

    }
}
