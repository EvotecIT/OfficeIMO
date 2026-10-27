using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed class WordPdfMixedInlineImageUnavailableTests {
        [Fact]
        public void InlineImageRetainsCropWarningAfterSuccessfulPreparation() {
            using WordDocument word = WordDocument.Create();
            WordParagraph paragraph = word.AddParagraph("Before");
            using var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2));
            WordImage image = paragraph.AddText("").AddImage(pixels, "image.png", 80, 40);
            image.CropLeft = -5000;
            paragraph.AddText("After");

            var result = word.ConvertToPdfWithEvidence();
            var warning = Assert.Single(result.Report.Warnings, warning => warning.Code == "native-image-crop-approximated");
            Assert.Equal("inline image", warning.Source);
            Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code is "native-body-image-unavailable" or "NativeBodyImageUnsupported");
            Assert.Single(PdfImageExtractor.ExtractImagePlacements(result.Value.ToBytes()));
        }

        [Theory]
        [InlineData("body")]
        [InlineData("cell")]
        [InlineData("header")]
        public void LinkedInlineImageLeavesTextAndAnUnavailableImageWarning(string story) {
            using WordDocument word = WordDocument.Create();
            word.AddParagraph("Body");
            if (story == "header") word.AddHeadersAndFooters();
            var paragraph = story == "cell" ? word.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] :
                story == "header" ? word.Sections[0].Header.Default!.AddParagraph() : word.AddParagraph();
            paragraph.Text = "Before";
            paragraph.AddText("").AddImage(new Uri("https://example.invalid/image.png"), 80, 80);
            paragraph.AddText("After");
            var result = word.ConvertToPdfWithEvidence();
            var warning = Assert.Single(result.Report.Warnings, warning => warning.Code == "native-body-image-unavailable");
            if (story == "body") Assert.Equal("body paragraph image", warning.Source);
            // Running headers are artifacts; include that paint when testing retained header content.
            var read = PdfReadDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }), new PdfLoadOptions { IncludeArtifactText = true });
            Assert.Contains("BeforeAfter", read.ExtractText());
            Assert.Empty(read.ExtractImages());
        }
    }
}
