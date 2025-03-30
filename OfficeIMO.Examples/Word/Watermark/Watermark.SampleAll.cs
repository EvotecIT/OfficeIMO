using System;
using System.IO;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        public static void Watermark_SampleAll(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with various watermarks");
            string filePath = Path.Combine(folderPath, "Document with Various Watermarks.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var imagePathToAdd = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Console.WriteLine("Initial Counts:");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Document Images: {document.Images.Count}");
                Console.WriteLine("---");

                // Section 0: Default Text Watermark
                document.AddParagraph("This is Section 0.");
                document.AddHeadersAndFooters(); // Ensure headers/footers exist
                var watermark1 = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "DRAFT");
                watermark1.Color = Color.Gray;

                Console.WriteLine("After adding Section 0 Default Text Watermark:");
                Console.WriteLine($"  Section 0 Watermarks: {document.Sections[0].Watermarks.Count}");
                Console.WriteLine($"  Section 0 Header Default Watermarks: {document.Sections[0].Header.Default.Watermarks.Count}");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Document Images: {document.Images.Count}"); // Text watermark isn't an Image
                Console.WriteLine("---");

                // Add Section 1 with different first page header
                document.AddSection();
                document.Sections[1].DifferentFirstPage = true; // This also adds header/footer parts if needed
                document.AddParagraph("This is the first page of Section 1.");
                var watermark2 = document.Sections[1].Header.First.AddWatermark(WordWatermarkStyle.Text, "FIRST PAGE S1");
                watermark2.Color = Color.LightBlue;

                document.AddPageBreak(); // Move to the default page of section 1

                document.AddParagraph("This is a default page of Section 1.");
                var watermark3 = document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd);

                Console.WriteLine("After adding Section 1 First Page (Text) and Default (Image) Watermarks:");
                Console.WriteLine($"  Section 0 Watermarks: {document.Sections[0].Watermarks.Count}");
                Console.WriteLine($"  Section 1 Watermarks: {document.Sections[1].Watermarks.Count}"); // Should be 2 (First + Default)
                Console.WriteLine($"  Section 1 Header First Watermarks: {document.Sections[1].Header.First.Watermarks.Count}");
                Console.WriteLine($"  Section 1 Header Default Watermarks: {document.Sections[1].Header.Default.Watermarks.Count}");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}"); // Should be 3 (S0.Default + S1.First + S1.Default)
                Console.WriteLine($"  Document Images: {document.Images.Count}"); // Should be 1 (from S1.Default)
                 Console.WriteLine($"  S1 Default Header Images: {document.Sections[1].Header.Default.Images.Count}"); // Should be 1
                Console.WriteLine("---");

                // Add Section 2 with different odd/even pages
                document.AddSection();
                document.Sections[2].DifferentOddAndEvenPages = true; // Adds header/footer parts if needed
                document.AddParagraph("This is an odd page of Section 2 (inherits Default header).");
                // Section 2 Default inherits from Section 1 Default (Image watermark)

                document.AddPageBreak();
                document.AddParagraph("This is an even page of Section 2.");
                var watermark4 = document.Sections[2].Header.Even.AddWatermark(WordWatermarkStyle.Text, "EVEN S2");
                watermark4.Color = Color.Orange;

                Console.WriteLine("After adding Section 2 Even Page Watermark (Default inherited):");
                Console.WriteLine($"  Section 0 Watermarks: {document.Sections[0].Watermarks.Count}");
                Console.WriteLine($"  Section 1 Watermarks: {document.Sections[1].Watermarks.Count}");
                Console.WriteLine($"  Section 2 Watermarks: {document.Sections[2].Watermarks.Count}"); // Should be 2 (Default(inherited from S1) + Even)
                Console.WriteLine($"  Section 2 Header Default Watermarks: {document.Sections[2].Header.Default?.Watermarks.Count ?? 0}"); // Might be 0 if not explicitly created yet? Check inheritance
                Console.WriteLine($"  Section 2 Header Even Watermarks: {document.Sections[2].Header.Even.Watermarks.Count}");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}"); // Should be 4 (S0.Def + S1.First + S1.Def + S2.Even)
                Console.WriteLine($"  Document Images: {document.Images.Count}"); // Still 1 (from S1.Default)
                Console.WriteLine("---");

                 // Test removing a watermark
                Console.WriteLine("Removing Section 1 First Page Watermark...");
                document.Sections[1].Header.First.Watermarks[0].Remove();
                Console.WriteLine($"  Section 1 Watermarks after remove: {document.Sections[1].Watermarks.Count}"); // Should be 1 (Default)
                Console.WriteLine($"  Document Watermarks after remove: {document.Watermarks.Count}"); // Should be 3
                 Console.WriteLine($"  S1 First Header Watermarks: {document.Sections[1].Header.First.Watermarks.Count}");
                 Console.WriteLine("---");

                document.Save(openWord);
            }
        }
    }
}