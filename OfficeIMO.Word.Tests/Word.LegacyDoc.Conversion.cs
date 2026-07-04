using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_Convert_DocxToDocAndBack_RoundTripsSupportedContent() {
            string docxPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string roundTripPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");

            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Conversion first paragraph");
                document.AddParagraph("Conversion second paragraph");
                document.Save(docxPath);
            }

            WordDocument.Convert(docxPath, docPath);

            AssertOleCompoundFile(docPath);
            using (WordDocument legacy = WordDocument.Load(docPath)) {
                Assert.True(legacy.WasLoadedFromLegacyDoc);
                Assert.Empty(legacy.LegacyDocUnsupportedFeatures);
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion first paragraph");
                Assert.Contains(legacy.Paragraphs, paragraph => paragraph.Text == "Conversion second paragraph");
            }

            WordDocument.Convert(docPath, roundTripPath);

            using WordDocument roundTrip = WordDocument.Load(roundTripPath);
            Assert.False(roundTrip.WasLoadedFromLegacyDoc);
            Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.Text == "Conversion first paragraph");
            Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.Text == "Conversion second paragraph");
        }

        [Fact]
        public void LegacyDoc_Convert_BlocksUnsupportedLegacyContentUnlessLossIsAllowed() {
            string docPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Preserve-only body"));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => WordDocument.Convert(docPath, blockedPath));

            Assert.Contains("unsupported or preserve-only", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(blockedPath));

            WordDocument.Convert(docPath, allowedPath, new WordDocumentConversionOptions {
                AllowLossyLegacyConversion = true
            });

            using WordDocument converted = WordDocument.Load(allowedPath);
            Assert.False(converted.WasLoadedFromLegacyDoc);
            Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text == "Preserve-only body");
        }

        private static void AssertOleCompoundFile(string path) {
            byte[] bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length > 8);
            Assert.Equal(0xd0, bytes[0]);
            Assert.Equal(0xcf, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xe0, bytes[3]);
        }
    }
}
