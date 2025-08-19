using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void WordImage_Clone_ReusesImagePart() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageClone.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph1 = document.AddParagraph();
            paragraph1.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var paragraph2 = document.AddParagraph();
            var clone = paragraph1.Image.Clone(paragraph2);

            Assert.Equal(paragraph1.Image.RelationshipId, clone.RelationshipId);
            Assert.Equal(2, document.Images.Count);
            Assert.NotNull(document._wordprocessingDocument.MainDocumentPart);
            Assert.Single(document._wordprocessingDocument.MainDocumentPart!.ImageParts);

            document.Save(false);
        }
    }
}
