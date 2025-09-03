using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ParagraphBordersAndShading() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphBordersAndShading.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Border and shading");
                paragraph.Borders.LeftStyle = BorderValues.Thick;
                paragraph.Borders.LeftColor = Color.Red;
                paragraph.Borders.LeftSize = 24;
                paragraph.ShadingFillColor = Color.LightGray;
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs[0];
                Assert.Equal(BorderValues.Thick, paragraph.Borders.LeftStyle);
                Assert.Equal(Color.Red.ToHexColor(), paragraph.Borders.LeftColor!.Value.ToHexColor());
                Assert.Equal(24U, paragraph.Borders.LeftSize!.Value);
                Assert.Equal(Color.LightGray.ToHexColor(), paragraph.ShadingFillColorHex);
            }
        }
    }
}
