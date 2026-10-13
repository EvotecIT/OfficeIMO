using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(0, WordVerticalCharacterAlignment.Top)]
        [InlineData(1, WordVerticalCharacterAlignment.Center)]
        [InlineData(2, WordVerticalCharacterAlignment.Baseline)]
        [InlineData(3, WordVerticalCharacterAlignment.Bottom)]
        [InlineData(4, WordVerticalCharacterAlignment.Auto)]
        public void LegacyDoc_LoadsNativeFontAlignmentValues(byte nativeValue, WordVerticalCharacterAlignment expected) {
            byte[] bytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphVerticalCharacterAlignment(nativeValue);
            using WordDocument document = WordDocument.Load(new MemoryStream(bytes));

            Assert.Equal(expected, Assert.Single(document.Paragraphs).VerticalCharacterAlignmentOnLine);
        }

        [Theory]
        [InlineData(0, WordVerticalCharacterAlignment.Top)]
        [InlineData(1, WordVerticalCharacterAlignment.Center)]
        [InlineData(2, WordVerticalCharacterAlignment.Baseline)]
        [InlineData(3, WordVerticalCharacterAlignment.Bottom)]
        [InlineData(4, WordVerticalCharacterAlignment.Auto)]
        public void LegacyDoc_WritesNativeFontAlignmentValues(byte nativeValue, WordVerticalCharacterAlignment alignment) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Aligned text").VerticalCharacterAlignmentOnLine = alignment;

            byte[] bytes = document.ToBytes(WordFileFormat.Doc);
            byte[] stream = ReadCompoundStream(bytes, "WordDocument");
            Assert.True(ContainsBytePattern(stream, 0x39, 0x44, nativeValue, 0x00),
                $"Expected sprmPWAlignFont to encode {alignment} as {nativeValue}.");
        }
    }
}
