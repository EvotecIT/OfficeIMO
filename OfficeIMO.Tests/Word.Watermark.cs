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
        private static WordSection GetSection(WordDocument document, int index) {
            Assert.InRange(index, 0, document.Sections.Count - 1);
            return document.Sections[index];
        }

        private static WordHeaders GetHeaders(WordDocument document, int sectionIndex) {
            var section = GetSection(document, sectionIndex);
            return Assert.IsType<WordHeaders>(section.Header);
        }

        private static WordHeaders GetDocumentHeaders(WordDocument document) {
            return Assert.IsType<WordHeaders>(document.Header);
        }

        private static WordHeader GetHeader(WordDocument document, int sectionIndex, HeaderFooterValues type) {
            var headers = GetHeaders(document, sectionIndex);
            if (type == HeaderFooterValues.Default) {
                return Assert.IsType<WordHeader>(headers.Default);
            }

            if (type == HeaderFooterValues.First) {
                return Assert.IsType<WordHeader>(headers.First);
            }

            if (type == HeaderFooterValues.Even) {
                return Assert.IsType<WordHeader>(headers.Even);
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, null);
        }

        private static WordHeader GetHeader(WordDocument document, HeaderFooterValues type) {
            var headers = GetDocumentHeaders(document);
            if (type == HeaderFooterValues.Default) {
                return Assert.IsType<WordHeader>(headers.Default);
            }

            if (type == HeaderFooterValues.First) {
                return Assert.IsType<WordHeader>(headers.First);
            }

            if (type == HeaderFooterValues.Even) {
                return Assert.IsType<WordHeader>(headers.Even);
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, null);
        }

        [Fact]
        public void Test_CreatingWordDocumentWithWatermark() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                var firstSection = GetSection(document, 0);
                firstSection.SetMargins(WordMargin.Normal);

                var firstHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                var watermark = firstHeader.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                var secondSection = GetSection(document, 1);
                secondSection.AddHeadersAndFooters();
                secondSection.Margins.Type = WordMargin.Narrow;

                Assert.False(watermark.Stroked);
                Assert.False(watermark.AllowInCell);
                Assert.Equal(90, watermark.Rotation);
                Assert.Equal(131.95, watermark.Height);
                Assert.Equal(527.85, watermark.Width);
                Assert.Equal("silver", watermark.ColorHex);
                Assert.Equal(Color.Silver, watermark.Color);
                Assert.Equal("Watermark", watermark.Text);

                var secondHeader = GetHeader(document, 1, HeaderFooterValues.Default);
                secondHeader.AddWatermark(WordWatermarkStyle.Text, "Draft");

                document.Settings.SetBackgroundColor(Color.Azure);

                Assert.Equal(2, document.Watermarks.Count);

                var documentDefaultHeaderWatermarks = GetHeader(document, HeaderFooterValues.Default).Watermarks;
                Assert.Single(documentDefaultHeaderWatermarks);

                Assert.Single(firstHeader.Watermarks);
                Assert.Single(secondHeader.Watermarks);

                document.AddSection();

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var firstHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                var firstSectionWatermark = Assert.Single(firstHeader.Watermarks);
                Assert.Equal(Color.Silver, firstSectionWatermark.Color);
                Assert.Equal("silver", firstSectionWatermark.ColorHex);
                Assert.Equal("Watermark", firstSectionWatermark.Text);
                Assert.Equal(131.95, firstSectionWatermark.Height);
                Assert.Equal(527.85, firstSectionWatermark.Width);
                Assert.Equal(90, firstSectionWatermark.Rotation);
                Assert.False(firstSectionWatermark.Stroked);
                Assert.False(firstSectionWatermark.AllowInCell);

                firstSectionWatermark.Stroked = true;

                var thirdSection = GetSection(document, 2);
                thirdSection.AddHeadersAndFooters();
                var thirdHeader = GetHeader(document, 2, HeaderFooterValues.Default);
                var thirdWatermark = thirdHeader.AddWatermark(WordWatermarkStyle.Text, "Check me");
                thirdWatermark.Rotation = 180;

                var section2Watermark = Assert.Single(thirdHeader.Watermarks);
                Assert.Equal(Color.Silver, section2Watermark.Color);
                Assert.Equal("silver", section2Watermark.ColorHex);
                Assert.Equal("Check me", section2Watermark.Text);
                Assert.Equal(131.95, section2Watermark.Height);
                Assert.Equal(527.85, section2Watermark.Width);
                Assert.Equal(180, section2Watermark.Rotation);
                Assert.False(section2Watermark.Stroked);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var firstHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                var firstSectionWatermark = Assert.Single(firstHeader.Watermarks);
                Assert.Equal(Color.Silver, firstSectionWatermark.Color);
                Assert.Equal("silver", firstSectionWatermark.ColorHex);
                Assert.Equal("Watermark", firstSectionWatermark.Text);
                Assert.Equal(131.95, firstSectionWatermark.Height);
                Assert.Equal(527.85, firstSectionWatermark.Width);
                Assert.Equal(90, firstSectionWatermark.Rotation);
                Assert.True(firstSectionWatermark.Stroked);
                Assert.False(firstSectionWatermark.AllowInCell);

                Assert.Equal(3, document.Watermarks.Count);

                var documentDefaultHeaderWatermarks = GetHeader(document, HeaderFooterValues.Default).Watermarks;
                Assert.Single(documentDefaultHeaderWatermarks);

                var secondHeader = GetHeader(document, 1, HeaderFooterValues.Default);
                Assert.Single(secondHeader.Watermarks);

                var thirdHeader = GetHeader(document, 2, HeaderFooterValues.Default);
                Assert.Single(thirdHeader.Watermarks);

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
                var firstHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                firstHeader.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                document.AddSection();
                var secondSection = GetSection(document, 1);
                secondSection.AddHeadersAndFooters();
                var secondHeader = GetHeader(document, 1, HeaderFooterValues.Default);
                secondHeader.AddWatermark(WordWatermarkStyle.Text, "Draft");
                document.Settings.SetBackgroundColor(Color.Azure);
                document.AddSection();

                Assert.Equal(2, document.Watermarks.Count);
                var documentDefaultHeaderWatermarks = GetHeader(document, HeaderFooterValues.Default).Watermarks;
                Assert.Single(documentDefaultHeaderWatermarks);
                Assert.Single(firstHeader.Watermarks);
                Assert.Single(secondHeader.Watermarks);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Watermarks.Count);
                var documentDefaultHeaderWatermarks = GetHeader(document, HeaderFooterValues.Default).Watermarks;
                Assert.Single(documentDefaultHeaderWatermarks);

                var firstHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                Assert.Single(firstHeader.Watermarks);

                var secondHeader = GetHeader(document, 1, HeaderFooterValues.Default);
                Assert.Single(secondHeader.Watermarks);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save();
            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithWatermark3() {
            // this test adding watermark without adding headers/footers first (watermark is added in the header.default)
            string filePath = Path.Combine(_directoryWithFiles, "Test_CreatingWordDocumentWithWatermark2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");

                Assert.Empty(document.Watermarks);

                var firstSection = GetSection(document, 0);
                firstSection.AddWatermark(WordWatermarkStyle.Text, "Confidential");

                document.AddPageBreak();
                document.AddPageBreak();

                Assert.Single(document.Watermarks);

                var section = document.AddSection();
                section.AddWatermark(WordWatermarkStyle.Text, "Second Mark");

                Assert.Equal(2, document.Watermarks.Count);

                document.AddPageBreak();
                document.AddPageBreak();

                var thirdSection = document.AddSection();

                Assert.Equal(2, document.Watermarks.Count);

                thirdSection.AddWatermark(WordWatermarkStyle.Text, "New");

                document.AddPageBreak();
                document.AddPageBreak();

                Assert.Equal(3, document.Watermarks.Count);
                Assert.Single(GetSection(document, 0).Watermarks);
                Assert.Single(GetSection(document, 1).Watermarks);
                Assert.Single(GetSection(document, 2).Watermarks);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(3, document.Watermarks.Count);
                Assert.Single(GetSection(document, 0).Watermarks);
                Assert.Single(GetSection(document, 1).Watermarks);
                Assert.Single(GetSection(document, 2).Watermarks);

                var watermarks = document.Watermarks;
                Assert.NotEmpty(watermarks);
                var watermark = watermarks[0];
                watermark.Remove();

                Assert.Equal(2, document.Watermarks.Count);
                Assert.Empty(GetSection(document, 0).Watermarks);
                Assert.Single(GetSection(document, 1).Watermarks);
                Assert.Single(GetSection(document, 2).Watermarks);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
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

                var defaultHeader = GetHeader(document, 0, HeaderFooterValues.Default);
                var firstHeader = GetHeader(document, 0, HeaderFooterValues.First);
                var evenHeader = GetHeader(document, 0, HeaderFooterValues.Even);

                defaultHeader.AddWatermark(WordWatermarkStyle.Text, "Default");
                firstHeader.AddWatermark(WordWatermarkStyle.Text, "First");
                evenHeader.AddWatermark(WordWatermarkStyle.Text, "Even");

                var sectionWatermarks = GetSection(document, 0).Watermarks;
                Assert.Equal(3, sectionWatermarks.Count);

                foreach (var watermark in sectionWatermarks.ToList()) {
                    watermark.Remove();
                }

                Assert.Empty(GetSection(document, 0).Watermarks);
                Assert.Empty(defaultHeader.Watermarks);
                Assert.Empty(firstHeader.Watermarks);
                Assert.Empty(evenHeader.Watermarks);
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