using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for background images.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_SetBackgroundImage() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentBackgroundImage.docx");
            string imagePath = Path.Combine(_directoryWithImages, "BackgroundImage.png");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Background.SetImage(imagePath, OfficeLength.Pixels(600), OfficeLength.Pixels(800));

                var background = document._wordprocessingDocument!.MainDocumentPart!.Document.DocumentBackground;
                Assert.NotNull(background);
                var originalExtent = background!.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>()!.Anchor!.Extent!;
                Assert.Equal(5715000L, originalExtent.Cx!.Value);
                Assert.Equal(7620000L, originalExtent.Cy!.Value);

                using var imageStream = File.OpenRead(imagePath);
                document.Background.SetImage(imageStream, "BackgroundImage.png", OfficeLength.Inches(6.25), OfficeLength.Points(600));
                Assert.True(imageStream.CanRead);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var background = document._wordprocessingDocument!.MainDocumentPart!.Document.DocumentBackground;
                Assert.NotNull(background);
                var savedExtent = Assert.Single(System.Xml.Linq.XElement.Parse(background!.OuterXml)
                    .Descendants(System.Xml.Linq.XName.Get("extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")));
                Assert.Equal("5715000", savedExtent.Attribute("cx")?.Value);
                Assert.Equal("7620000", savedExtent.Attribute("cy")?.Value);
            }
        }
    }
}
