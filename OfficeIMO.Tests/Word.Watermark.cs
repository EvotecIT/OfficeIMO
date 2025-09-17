using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using Xunit;
using Color = SixLabors.ImageSharp.Color;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithWatermark() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                document.Sections[0].SetMargins(WordMargin.Normal);

                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");
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

                document.Sections[1].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Draft");

                document.Settings.SetBackgroundColor(Color.Azure);

                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header!.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header!.Default.Watermarks.Count == 1);

                document.AddSection();

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {

                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Text == "Watermark");
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Rotation == 90);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Stroked == false);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].AllowInCell == false);

                document.Sections[0].Header!.Default.Watermarks[0].Stroked = true;

                // let's add first headers and footers to section 2 so we can add watermark to it
                document.Sections[2].AddHeadersAndFooters();
                var watermark = document.Sections[2].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Check me");
                watermark.Rotation = 180;

                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Text == "Check me");
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Rotation == 180);
                Assert.True(document.Sections[2].Header!.Default.Watermarks[0].Stroked == false);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx"))) {

                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Color == Color.Silver);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].ColorHex == "silver");
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Text == "Watermark");
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Height == 131.95);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Width == 527.85);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Rotation == 90);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].Stroked == true);
                Assert.True(document.Sections[0].Header!.Default.Watermarks[0].AllowInCell == false);

                Assert.True(document.Watermarks.Count == 3);
                Assert.True(document.Header!.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header!.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[2].Header!.Default.Watermarks.Count == 1);

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
                document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Draft");
                document.Settings.SetBackgroundColor(Color.Azure);
                document.AddSection();

                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header!.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header!.Default.Watermarks.Count == 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
                Assert.True(document.Watermarks.Count == 2);
                Assert.True(document.Header!.Default.Watermarks.Count == 1); // this is actually first section's header.default
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 1);
                Assert.True(document.Sections[1].Header!.Default.Watermarks.Count == 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx"))) {
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
        public void Test_RemoveWatermarkFromAllHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_RemoveWatermarkFromHeaders.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Default");
                document.Sections[0].Header!.First.AddWatermark(WordWatermarkStyle.Text, "First");
                document.Sections[0].Header!.Even.AddWatermark(WordWatermarkStyle.Text, "Even");

                Assert.True(document.Sections[0].Watermarks.Count == 3);

                foreach (var watermark in document.Sections[0].Watermarks.ToList()) {
                    watermark.Remove();
                }

                Assert.True(document.Sections[0].Watermarks.Count == 0);
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 0);
                Assert.True(document.Sections[0].Header!.First.Watermarks.Count == 0);
                Assert.True(document.Sections[0].Header!.Even.Watermarks.Count == 0);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImageWatermark() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithImageWatermark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.AddHeadersAndFooters();
                var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Image, imagePath);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Watermarks.Count == 1);
                Assert.True(document.Sections[0].Header!.Default.Watermarks.Count == 1);
            }
        }

        [Fact]
        public void Test_WatermarkImageDimensions() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkImageDimensions.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.AddHeadersAndFooters();
                var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Image, imagePath);
                Assert.True(watermark.Width > 0);
                Assert.True(watermark.Height > 0);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var watermark = document.Sections[0].Header!.Default.Watermarks[0];
                Assert.True(watermark.Width > 0);
                Assert.True(watermark.Height > 0);
            }
        }

        [Fact]
        public void Test_WatermarkOffsetsAndScale() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkOffsetsAndScale.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test");
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Offset", 10, 20, 2.0);
                Assert.Equal(10, watermark.HorizontalOffset);
                Assert.Equal(20, watermark.VerticalOffset);
                Assert.True(watermark.Width > 0);
                Assert.True(watermark.Height > 0);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var watermark = document.Sections[0].Header!.Default.Watermarks[0];
                Assert.Equal(10, watermark.HorizontalOffset);
                Assert.Equal(20, watermark.VerticalOffset);
                Assert.True(watermark.Width > 0);
                Assert.True(watermark.Height > 0);
            }
        }

        [Fact]
        public void Test_WatermarkColorSupportsHex() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkColorSupportsHex.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Hex");
                watermark.Color = Color.Red;
                document.Save();
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var fill = wordDoc.MainDocumentPart!.HeaderParts.First().Header!.Descendants<V.Shape>().First().FillColor?.Value;
                Assert.True(fill == "#ff0000");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var watermark = document.Watermarks[0];
                Assert.True(watermark.ColorHex == "ff0000");
                Assert.True(watermark.Color == Color.Red);
            }
        }

        [Fact]
        public void Test_WatermarkColorSupportsUppercaseHexWithHash() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkColorSupportsUppercaseHexWithHash.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Upper");
                watermark.ColorHex = "#FF00FF";
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var watermark = document.Watermarks[0];
                Assert.Equal("ff00ff", watermark.ColorHex);
                Assert.Equal(Color.Magenta, watermark.Color);
            }
        }

        [Fact]
        public void Test_WatermarkSupportsMultipleColorInputs() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkSupportsMultipleColorInputs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                // SixLabors colors
                var red = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Red");
                red.Color = Color.Red;

                document.AddSection();
                document.Sections[1].AddHeadersAndFooters();
                var green = document.Sections[1].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Green");
                green.Color = Color.Green;

                document.AddSection();
                document.Sections[2].AddHeadersAndFooters();
                var blue = document.Sections[2].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Blue");
                blue.Color = Color.Blue;

                // Hex without '#'
                document.AddSection();
                document.Sections[3].AddHeadersAndFooters();
                var magenta = document.Sections[3].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Magenta");
                magenta.ColorHex = "ff00ff";

                // Hex with '#'
                document.AddSection();
                document.Sections[4].AddHeadersAndFooters();
                var cyan = document.Sections[4].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Cyan");
                cyan.ColorHex = "#00ffff";

                // Named color string
                document.AddSection();
                document.Sections[5].AddHeadersAndFooters();
                var yellow = document.Sections[5].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Yellow");
                yellow.ColorHex = "yellow";

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("ff0000", document.Sections[0].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Red, document.Sections[0].Header!.Default.Watermarks[0].Color);

                Assert.Equal("008000", document.Sections[1].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Green, document.Sections[1].Header!.Default.Watermarks[0].Color);

                Assert.Equal("0000ff", document.Sections[2].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Blue, document.Sections[2].Header!.Default.Watermarks[0].Color);

                Assert.Equal("ff00ff", document.Sections[3].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Magenta, document.Sections[3].Header!.Default.Watermarks[0].Color);

                Assert.Equal("00ffff", document.Sections[4].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Cyan, document.Sections[4].Header!.Default.Watermarks[0].Color);

                Assert.Equal("ffff00", document.Sections[5].Header!.Default.Watermarks[0].ColorHex);
                Assert.Equal(Color.Yellow, document.Sections[5].Header!.Default.Watermarks[0].Color);
            }
        }

        [Theory]
        [InlineData("red", "ff0000")]
        [InlineData("#00FF00", "00ff00")]
        [InlineData("0000ff", "0000ff")]
        [InlineData("#ABC", "aabbcc")]
        [InlineData("abc", "aabbcc")]
        public void Test_WatermarkColorRoundTripAndRendering(string input, string expectedHex) {
            string filePath = Path.Combine(_directoryWithFiles, $"Test_WatermarkColorRoundTripAndRendering_{expectedHex}.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Color");
                watermark.ColorHex = input;
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var watermark = document.Sections[0].Header!.Default.Watermarks[0];
                Assert.Equal(expectedHex, watermark.ColorHex);
                Assert.Equal(Color.Parse(expectedHex), watermark.Color);
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var headerPart = wordDoc.MainDocumentPart!.HeaderParts.First();
                var shape = headerPart.Header!.Descendants<V.Shape>().First();
                var fill = shape.GetFirstChild<V.Fill>();
                var textPath = shape.GetFirstChild<V.TextPath>();

                Assert.Equal("#" + expectedHex, shape.FillColor?.Value);
                Assert.Equal("#" + expectedHex, fill?.Color);
                var fillAttr = textPath?.GetAttribute("fillcolor", string.Empty).Value;
                var strokeAttr = textPath?.GetAttribute("strokecolor", string.Empty).Value;
                Assert.Equal("#" + expectedHex, fillAttr);
                Assert.Equal("#" + expectedHex, strokeAttr);
            }
        }

        [Fact]
        public void Test_WatermarkInvalidColorThrows() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkInvalidColorThrows.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Invalid");
                Assert.Throws<ArgumentException>(() => watermark.ColorHex = "notacolor");
            }
        }

        [Fact]
        public void Test_WatermarkEmptyColorThrows() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_WatermarkEmptyColorThrows.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var watermark = document.Sections[0].Header!.Default.AddWatermark(WordWatermarkStyle.Text, "Invalid");
                Assert.Throws<ArgumentException>(() => watermark.ColorHex = "");
            }
        }
    }
}