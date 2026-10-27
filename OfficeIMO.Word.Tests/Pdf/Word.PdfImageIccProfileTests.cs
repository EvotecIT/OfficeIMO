using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed class WordPdfImageIccProfileTests {
        [Fact]
        public void InlinePngKeepsAProfileWhoseTagTypesHaveNonZeroReservedBytes() {
            byte[] profile = PdfPngTestImages.MarkIccTagTypeReservedBytes(PdfIccProfiles.SrgbIec6196621);
            using WordDocument word = WordDocument.Create();
            using var pixels = new MemoryStream(PdfPngTestImages.CreateRgbaPngWithIccProfile(profile));
            word.AddParagraph("Before").AddText("").InsertImage(pixels, "diagram.png", 80, 40);

            var result = word.ConvertToPdfWithEvidence();
            Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeBodyImageUnsupported");
            byte[] pdf = result.Value.ToBytes();
            Assert.Single(PdfImageExtractor.ExtractImagePlacements(pdf));
            var objects = PdfSyntax.ParseObjects(pdf).Map;
            PdfStream image = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
                stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image" &&
                    stream.Dictionary.Items.ContainsKey("SMask"));
            PdfArray colorSpace = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
            Assert.Equal("ICCBased", Assert.IsType<PdfName>(colorSpace.Items[0]).Name);
            PdfReference profileRef = Assert.IsType<PdfReference>(colorSpace.Items[1]);
            PdfStream embeddedProfile = Assert.IsType<PdfStream>(objects[profileRef.ObjectNumber].Value);
            Assert.Equal(3, embeddedProfile.Dictionary.Get<PdfNumber>("N")!.Value);
            Assert.Equal(profile, StreamDecoder.Decode(embeddedProfile.Dictionary, embeddedProfile.Data, objects));
        }
    }
}
