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
                document.Background.SetImage(imagePath, 600, 800);

                var background = document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground;
                Assert.NotNull(background);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var background = document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground;
                Assert.NotNull(background);
            }
        }
    }
}
