using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithWatermark() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                document.Sections[0].SetMargins(WordMargin.Normal);

                var watermark = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");
                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Margins.Type = WordMargin.Narrow;

                Assert.True(watermark.Stroked == false);
                Assert.True(watermark.AllowInCell == false);
                Assert.True(watermark.Rotation == 90);

                Assert.True(watermark.Height == 131.95, "Value was " + watermark.Height + " but should be " + "131.95");
                Assert.True(watermark.Width == 527.85);
                Assert.True(watermark.ColorHex == "silver");
                Assert.True(watermark.Color == Color.Silver);
                Assert.True(watermark.Text == "Watermark");

                document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Draft");

                document.Settings.SetBackgroundColor(Color.Azure);

                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header.Default.Watermarks.Count == 1);

                document.AddSection();

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {

                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Text == "Watermark");
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Rotation == 90);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Stroked == false);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].AllowInCell == false);

                document.Sections[0].Header.Default.Watermarks[0].Stroked = true;

                // let's add first headers and footers to section 2 so we can add watermark to it
                document.Sections[2].AddHeadersAndFooters();
                var watermark = document.Sections[2].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Check me");
                watermark.Rotation = 180;

                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Text == "Check me");
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Rotation == 180);
                Assert.True(document.Sections[2].Header.Default.Watermarks[0].Stroked == false);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {

                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Text == "Watermark");
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Rotation == 90);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].Stroked == true);
                Assert.True(document.Sections[0].Header.Default.Watermarks[0].AllowInCell == false);

                Assert.True(document.Watermarks.Count == 3);
                Assert.True(document.Header.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[2].Header.Default.Watermarks.Count == 1);

                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithWatermark2() {
            // this test adding watermark without adding headers/footers first (watermark is added in the header.default)
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");

                document.AddHeadersAndFooters();
                document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Draft");
                document.Settings.SetBackgroundColor(Color.Azure);
                document.AddSection();

                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header.Default.Watermarks.Count == 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header.Default.Watermarks.Count == 1);

                document.Save();
            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithWatermark3() {
            // this test adding watermark without adding headers/footers first (watermark is added in the header.default)
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");

                Assert.True(document.Watermarks.Count == 0);

                document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Confidential");

                document.AddPageBreak();
                document.AddPageBreak();

                Assert.True(document.Watermarks.Count == 1);

                var section = document.AddSection();
                section.AddWatermark(WordWatermarkStyle.Text, "Second Mark");

                Assert.True(document.Watermarks.Count == 2);

                document.AddPageBreak();
                document.AddPageBreak();

                var section1 = document.AddSection();

                Assert.True(document.Watermarks.Count == 2);

                document.Sections[2].AddWatermark(WordWatermarkStyle.Text, "New");

                document.AddPageBreak();
                document.AddPageBreak();

                Assert.True(document.Watermarks.Count == 3);
                Assert.True(document.Sections[0].Watermarks.Count == 1);
                Assert.True(document.Sections[1].Watermarks.Count == 1);
                Assert.True(document.Sections[2].Watermarks.Count == 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
                Assert.True(document.Watermarks.Count == 3);
                Assert.True(document.Sections[0].Watermarks.Count == 1);
                Assert.True(document.Sections[1].Watermarks.Count == 1);
                Assert.True(document.Sections[2].Watermarks.Count == 1);

                document.Watermarks[0].Remove();

                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Sections[0].Watermarks.Count == 0);
                Assert.True(document.Sections[1].Watermarks.Count == 1);
                Assert.True(document.Sections[2].Watermarks.Count == 1);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
                document.Save();
            }
        }

        [Fact]
        public void Test_AddRemoveMultipleWatermarks() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_AddRemoveMultipleWatermarks.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var imagePathToAdd = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // Initial State
                Assert.Equal(0, document.Watermarks.Count);
                document.AddParagraph("Section 0 Content");

                // 1. Add Text Watermark to S0 Default Header
                document.AddHeadersAndFooters(); // Ensure headers exist for S0
                var watermarkText0 = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "TEXT_S0_DEFAULT");
                watermarkText0.Color = Color.Blue;

                Assert.Equal(1, document.Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Header.Default.Watermarks.Count);

                // 2. Add Image Watermark to S1 Default Header
                document.AddSection();
                document.Sections[1].AddParagraph("Section 1 Content");
                document.Sections[1].AddHeadersAndFooters(); // Ensure headers exist for S1
                var watermarkImage1 = document.Sections[1].Header.Default.AddWatermark(WordWatermarkStyle.Image, imagePathToAdd);

                Assert.Equal(2, document.Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Watermarks.Count); // S0 unchanged
                Assert.Equal(1, document.Sections[0].Header.Default.Watermarks.Count);
                Assert.Equal(1, document.Sections[1].Watermarks.Count); // S1 has 1 definition
                Assert.Equal(1, document.Sections[1].Header.Default.Watermarks.Count);
                Assert.Null(document.Sections[1].Header.First); // First shouldn't exist yet

                // 3. Add Text Watermark to S1 First Header
                document.Sections[1].DifferentFirstPage = true;
                // AddHeadersAndFooters might be needed again AFTER setting DifferentFirstPage if it doesn't create parts automatically
                // Or ensure AddWatermark creates the header if needed (current logic does)
                var watermarkText1First = document.Sections[1].Header.First.AddWatermark(WordWatermarkStyle.Text, "TEXT_S1_FIRST");
                watermarkText1First.Color = Color.Orange;

                Assert.Equal(3, document.Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Watermarks.Count); // S0 unchanged
                Assert.Equal(2, document.Sections[1].Watermarks.Count); // S1 now has 2 definitions (Default Image, First Text)
                Assert.Equal(1, document.Sections[1].Header.Default.Watermarks.Count);
                Assert.NotNull(document.Sections[1].Header.First);
                Assert.Equal(1, document.Sections[1].Header.First.Watermarks.Count);

                document.Save();
            }

            // Re-load and verify counts
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(3, document.Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Header.Default.Watermarks.Count);
                Assert.Equal(2, document.Sections[1].Watermarks.Count);
                Assert.Equal(1, document.Sections[1].Header.Default.Watermarks.Count);
                Assert.NotNull(document.Sections[1].Header.First);
                Assert.Equal(1, document.Sections[1].Header.First.Watermarks.Count);

                // Verify content
                Assert.Equal("TEXT_S0_DEFAULT", document.Sections[0].Header.Default.Watermarks[0].Text);
                Assert.Equal(Color.Blue, document.Sections[0].Header.Default.Watermarks[0].Color);
                Assert.Equal(WordWatermarkStyle.Image, document.Sections[1].Header.Default.Watermarks[0].Type);
                Assert.Equal("TEXT_S1_FIRST", document.Sections[1].Header.First.Watermarks[0].Text);
                Assert.Equal(Color.Orange, document.Sections[1].Header.First.Watermarks[0].Color);

                // 4. Remove Image Watermark (S1 Default)
                var imageToRemove = document.Sections[1].Header.Default.Watermarks.FirstOrDefault(wm => wm.Type == WordWatermarkStyle.Image);
                Assert.NotNull(imageToRemove);
                imageToRemove.Remove();

                Assert.Equal(2, document.Watermarks.Count);
                Assert.Equal(1, document.Sections[0].Watermarks.Count); // S0 unchanged
                Assert.Equal(1, document.Sections[1].Watermarks.Count); // S1 should have 1 left (First Text)
                Assert.Equal(0, document.Sections[1].Header.Default.Watermarks.Count); // S1 Default should be empty
                Assert.Equal(1, document.Sections[1].Header.First.Watermarks.Count); // S1 First unchanged

                // 5. Remove Text Watermark (S0 Default)
                var textToRemove = document.Sections[0].Header.Default.Watermarks.FirstOrDefault(wm => wm.Type == WordWatermarkStyle.Text);
                Assert.NotNull(textToRemove);
                textToRemove.Remove();

                Assert.Equal(1, document.Watermarks.Count); // Only S1 First Text should remain
                Assert.Equal(0, document.Sections[0].Watermarks.Count);
                Assert.Equal(0, document.Sections[0].Header.Default.Watermarks.Count);
                Assert.Equal(1, document.Sections[1].Watermarks.Count);
                Assert.Equal(0, document.Sections[1].Header.Default.Watermarks.Count);
                Assert.Equal(1, document.Sections[1].Header.First.Watermarks.Count);

                document.Save();
            }

            // Final Load check
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.Watermarks.Count);
                Assert.Equal(0, document.Sections[0].Watermarks.Count);
                Assert.Equal(1, document.Sections[1].Watermarks.Count);
                Assert.Equal(0, document.Sections[1].Header.Default.Watermarks.Count);
                Assert.NotNull(document.Sections[1].Header.First);
                Assert.Equal(1, document.Sections[1].Header.First.Watermarks.Count);
                Assert.Equal("TEXT_S1_FIRST", document.Sections[1].Header.First.Watermarks[0].Text);
            }
        }
    }
}
