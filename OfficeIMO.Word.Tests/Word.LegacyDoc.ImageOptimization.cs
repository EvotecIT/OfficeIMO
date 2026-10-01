using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_WordProducedInlineCropAndDefaultFontSurviveImageOptimization() {
            using WordDocument document = WordDocument.Load(GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComCroppedInlinePicture.doc")));
            Assert.Empty(document.LegacyDocUnsupportedFeatures);
            var picture = Assert.Single(document.Images);
            int? crop = picture.CropLeft;
            Assert.True(crop > 0);
            AssertNormalFont(document);
            var report = document.OptimizeImages(new WordImageOptimizationOptions {
                Mode = OfficeImageOptimizationMode.DownsampleAndRecompress, TargetDpi = 144, JpegQuality = 65
            });
            Assert.True(report.BytesSaved > 0);
            using var output = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
            using WordDocument reopened = WordDocument.Load(output);
            AssertNormalFont(reopened);
            Assert.Contains("Independent Word producer - cropped image optimization.", reopened.Paragraphs.Select(p => p.Text));
            Assert.Equal(crop, Assert.Single(reopened.Images).CropLeft);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static void AssertNormalFont(WordDocument document) {
            var normal = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
                style => style.StyleId?.Value == "Normal");
            Assert.Equal("Aptos", normal.StyleRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
            Assert.Equal("24", normal.StyleRunProperties?.GetFirstChild<FontSize>()?.Val?.Value);
        }
    }
}
