using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed class PdfImageIccProfileTests {
        [Fact]
        public void InlinePngKeepsAProfileWhoseTagTypesHaveNonZeroReservedBytes() {
            byte[] profile = PdfPngTestImages.MarkIccTagTypeReservedBytes(PdfIccProfiles.SrgbIec6196621);
            using WordDocument word = WordDocument.Create();
            using var pixels = new MemoryStream(PdfPngTestImages.CreateRgbaPngWithIccProfile(profile));
            word.AddParagraph("Before").AddText("").InsertImage(pixels, "diagram.png", 80, 40);

            var result = word.ToPdfDocumentResult();
            Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeBodyImageUnsupported");
            Assert.Single(PdfImageExtractor.ExtractImagePlacements(result.Value.ToBytes()));
        }
    }
}
