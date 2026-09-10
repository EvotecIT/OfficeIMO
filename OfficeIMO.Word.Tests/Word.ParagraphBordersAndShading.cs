using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ParagraphBordersAddedAfterShadingAndSpacingRemainSchemaValid(bool usePreset) {
            using var document = WordDocument.Create();
            var paragraph = document.AddParagraph("Border after existing paragraph formatting");
            paragraph.ShadingFillColor = Color.LightGray;
            paragraph.LineSpacingAfterPoints = 12;
            if (usePreset) {
                paragraph.Borders.Type = WordBorder.Box;
                paragraph.Borders.Type = WordBorder.Shadow;
            } else {
                paragraph.Borders.BottomStyle = WordBorderStyle.Single;
                paragraph.Borders.RightStyle = WordBorderStyle.Single;
                paragraph.Borders.TopStyle = WordBorderStyle.Single;
                paragraph.Borders.LeftStyle = WordBorderStyle.Single;
            }

            using var stream = document.ToStream();
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
            var properties = package.MainDocumentPart!.Document.Body!.GetFirstChild<Paragraph>()!.ParagraphProperties!;
            Assert.Single(properties.Elements<ParagraphBorders>());
            Assert.Equal(Color.LightGray.ToRgbHex(), properties.Shading!.Fill!.Value);
            Assert.Equal("240", properties.SpacingBetweenLines!.After!.Value);
        }

        [Fact]
        public void Test_ParagraphBordersAndShading() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphBordersAndShading.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Border and shading");
                paragraph.Borders.LeftStyle = WordBorderStyle.Thick;
                paragraph.Borders.LeftColor = Color.Red;
                paragraph.Borders.LeftSize = 24;
                paragraph.ShadingFillColor = Color.LightGray;
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs[0];
                Assert.Equal(WordBorderStyle.Thick, paragraph.Borders.LeftStyle);
                Assert.Equal(Color.Red.ToRgbHex(), paragraph.Borders.LeftColor!.Value.ToRgbHex());
                Assert.Equal(24U, paragraph.Borders.LeftSize);
                Assert.Equal(Color.LightGray.ToRgbHex(), paragraph.ShadingFillColorHex);
            }
        }
    }
}
