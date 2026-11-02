using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

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

                var section0 = document.Sections[0];
                var section0Header = GetRequiredHeader(section0);
                section0Header.AddParagraph("Section 0 - In header");
                section0.PageSetup.Margins.Type = WordMargin.Normal;

                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Left.GetValueOrDefault().ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Right.GetValueOrDefault().ToTwips())));

                Console.WriteLine(document.Sections[0].PageSetup.Margins.Type);

                document.Sections[0].PageSetup.Margins.Type = WordMargin.Wide;

                Console.WriteLine(document.Sections[0].PageSetup.Margins.Type);

                Console.WriteLine("----");
                var watermark = section0Header.AddWatermark(WordWatermarkStyle.Text, "Watermark");
                watermark.Color = Color.Red;

                Console.WriteLine(watermark.Color);

                Console.WriteLine(watermark.Rotation);

                watermark.Rotation = 180;

                Console.WriteLine(watermark.Rotation);

                watermark.Stroked = true;

                Console.WriteLine(watermark.Height);
                Console.WriteLine(watermark.Width);

                // width and height in points (HTML wise)
                watermark.Height = global::OfficeIMO.OfficeLength.Points(100.15);
                watermark.Width = global::OfficeIMO.OfficeLength.Points(500.18);

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddSection();

                document.AddParagraph("Section 1");

                var section1 = document.Sections[1];
                var section1Header = GetRequiredHeader(section1);
                section1Header.AddParagraph("Section 1 - In header");
                section1.PageSetup.Margins.Type = WordMargin.Narrow;
                Console.WriteLine("----");

                Console.WriteLine("Section 0 - Paragraphs Count: " + section0Header.Paragraphs.Count);
                Console.WriteLine("Section 1 - Paragraphs Count: " + section1Header.Paragraphs.Count);

                Console.WriteLine("----");
                section1.AddParagraph("Test");
                section1Header.AddWatermark(WordWatermarkStyle.Text, "Draft");

                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Left.GetValueOrDefault().ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Right.GetValueOrDefault().ToTwips())));

                Console.WriteLine(checked((uint)(document.Sections[1].PageSetup.Margins.Left.GetValueOrDefault().ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[1].PageSetup.Margins.Right.GetValueOrDefault().ToTwips())));

                Console.WriteLine(document.Sections[1].PageSetup.Margins.Type);


                document.Settings.SetBackgroundColor(Color.Azure);

                Console.WriteLine("----");

                var documentDefaultHeader = document.PageSetup.GetHeader(WordHeaderFooterType.Default) ?? throw new InvalidOperationException("The default document header must exist when inspecting watermarks.");
                Console.WriteLine("Watermarks in default header: " + documentDefaultHeader.Watermarks.Count);

                var documentDefaultFooter = document.PageSetup.GetFooter(WordHeaderFooterType.Default) ?? throw new InvalidOperationException("The default document footer must exist when inspecting watermarks.");
                Console.WriteLine("Watermarks in default footer: " + documentDefaultFooter.Watermarks.Count);

                Console.WriteLine("Watermarks in section 0: " + document.Sections[0].Watermarks.Count);
                Console.WriteLine("Watermarks in section 0 (header): " + section0Header.Watermarks.Count);
                Console.WriteLine("Paragraphs in section 0 (header): " + section0Header.Paragraphs.Count);

                Console.WriteLine("Watermarks in section 1: " + document.Sections[1].Watermarks.Count);
                Console.WriteLine("Watermarks in section 1 (header): " + section1Header.Watermarks.Count);
                Console.WriteLine("Paragraphs in section 1 (header): " + section1Header.Paragraphs.Count);

                Console.WriteLine("Watermarks in document: " + document.Watermarks.Count);

                document.Save();
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

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}
