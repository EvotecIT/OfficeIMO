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
                Console.WriteLine($"  Section 0 Header Default Watermarks: {document.Sections[0].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Document Images: {document.Images.Count}"); // Text watermark isn't an Image
                Console.WriteLine("---");

                // Section 1: Different First Page and Default Image Watermark
                document.AddSection();
                document.Sections[1].AddParagraph("This is Section 1 (Page 2 - uses Default Header).");
                document.AddPageBreak(); // Go to Page 3
                document.Sections[1].AddParagraph("This is Section 1 (Page 3 - uses Default Header).");

                document.Sections[1].DifferentFirstPage = true; // Enable different first page *for section 1*
                document.Sections[1].AddHeadersAndFooters(); // Ensure parts exist
                var watermark2 = document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd); // Added to Default Header of Section 1

                Console.WriteLine("After adding Section 1 Default Image Watermark & enabling First Page:");
                Console.WriteLine($"  Section 1 Watermarks: {document.Sections[1].Watermarks.Count}");
                Console.WriteLine($"  Section 1 Header Default Watermarks: {document.Sections[1].Header.Default?.Watermarks.Count ?? 0}");
                // Use null-conditional access for potentially non-existent First/Even headers
                Console.WriteLine($"  Section 1 First Header Watermarks: {document.Sections[1].Header.First?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 1 Even Header Watermarks: {document.Sections[1].Header.Even?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 1 Images: {document.Sections[1].Images.Count}"); // Image watermark is VML, not WordImage
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}"); // Should reflect total unique definitions
                Console.WriteLine($"  Document Images: {document.Images.Count}");
                Console.WriteLine("---");

                // Section 2: Different First Page, Add First Page Text Watermark
                document.AddSection();
                document.Sections[2].AddParagraph("This is Section 2 (Page 4 - uses First Page Header).");
                document.AddPageBreak(); // Go to Page 5
                document.Sections[2].AddParagraph("This is Section 2 (Page 5 - uses Default Header, inherited from Section 1).");

                document.Sections[2].DifferentFirstPage = true; // Enable different first page *for section 2*
                document.Sections[2].AddHeadersAndFooters(); // Ensure parts exist
                var watermark3 = document.Sections[2].Header.First.AddWatermark(WordWatermarkStyle.Text, "FINAL"); // Add to First Page Header of Section 2
                watermark3.Color = Color.Green;

                Console.WriteLine("After adding Section 2 First Page Text Watermark:");
                Console.WriteLine($"  Section 2 Watermarks: {document.Sections[2].Watermarks.Count}");
                Console.WriteLine($"  Section 2 Header Default Watermarks: {document.Sections[2].Header.Default?.Watermarks.Count ?? 0}"); // Likely 0 unless explicitly added
                Console.WriteLine($"  Section 2 First Header Watermarks: {document.Sections[2].Header.First?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 2 Even Header Watermarks: {document.Sections[2].Header.Even?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Document Watermarks: {document.Watermarks.Count}"); // Should reflect total unique definitions
                Console.WriteLine($"  Document Images: {document.Images.Count}");
                Console.WriteLine("---");


                Console.WriteLine("Final Watermark Counts by Definition:");
                Console.WriteLine($"  Total Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Sec 0 Default: {document.Sections[0].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Sec 1 Default: {document.Sections[1].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Sec 1 First:   {document.Sections[1].Header.First?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Sec 2 Default: {document.Sections[2].Header.Default?.Watermarks.Count ?? 0}"); // Note: Sec 2 Default Header *might not exist* if only First was added
                Console.WriteLine($"  Sec 2 First:   {document.Sections[2].Header.First?.Watermarks.Count ?? 0}");
                Console.WriteLine("---");



                Console.WriteLine("Before Removing Section 1 Default Watermark:");
                Console.WriteLine($"  Total Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Section 0 Watermarks: {document.Sections[0].Watermarks.Count}");
                Console.WriteLine($"  Section 0 Header Default Watermarks: {document.Sections[0].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 1 Watermarks: {document.Sections[1].Watermarks.Count}");
                Console.WriteLine($"  Section 1 Header Default Watermarks: {document.Sections[1].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 2 Watermarks: {document.Sections[2].Watermarks.Count}");
                Console.WriteLine($"  Section 2 Header Default Watermarks: {document.Sections[2].Header.Default?.Watermarks.Count ?? 0}");

                Console.WriteLine("Removing Section 1 Default Watermark (Image)");

                // Need to ensure we get the correct watermark object to remove
                var watermarkToRemove = document.Sections[1].Header.Default?.Watermarks.FirstOrDefault();
                if (watermarkToRemove != null) {
                    watermarkToRemove.Remove();
                }

                Console.WriteLine("After Removing Section 1 Default Watermark:");
                Console.WriteLine($"  Total Document Watermarks: {document.Watermarks.Count}");
                Console.WriteLine($"  Section 0 Watermarks: {document.Sections[0].Watermarks.Count}");
                Console.WriteLine($"  Section 0 Header Default Watermarks: {document.Sections[0].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 1 Watermarks: {document.Sections[1].Watermarks.Count}");
                Console.WriteLine($"  Section 1 Header Default Watermarks: {document.Sections[1].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine($"  Section 2 Watermarks: {document.Sections[2].Watermarks.Count}");
                Console.WriteLine($"  Section 2 Header Default Watermarks: {document.Sections[2].Header.Default?.Watermarks.Count ?? 0}");
                Console.WriteLine("---");


                document.Save(openWord);
            }
        }
    }
}