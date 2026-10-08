using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeListDefinitions_TruncatedLevelPropertyOperandReportsNumberingLoss(bool paragraphFormatting) {
        using WordDocument source = CreateNativeListDefinitionControl();
        Level level = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().Elements<Level>().First();
        level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Bold());
        level.PreviousParagraphProperties = new PreviousParagraphProperties(new ContextualSpacing { Val = true });
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"], table = (byte[])compound.Streams["1Table"].Clone();
        int firstLevel = BitConverter.ToInt32(word, 0x2E2) + BitConverter.ToInt32(word, 0x2E6);
        int offset = firstLevel + 28 + (paragraphFormatting ? 0 : table[firstLevel + 25]);
        Assert.Equal(3, table[firstLevel + (paragraphFormatting ? 25 : 24)]);
        // Both property buffers have a three-byte bool SPRM. Replace it with a
        // two-byte operand SPRM, leaving only one byte inside the declared buffer.
        table[offset] = paragraphFormatting ? (byte)0x5E : (byte)0x43;
        table[offset + 1] = 0x4A;
        byte[] altered = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["1Table"] = table });
        using WordDocument reopened = WordDocument.Load(new MemoryStream(altered));
        Assert.Contains(reopened.LegacyDocUnsupportedFeatures, feature => feature.Kind == LegacyDocUnsupportedFeatureKind.Numbering);
        Assert.Throws<NotSupportedException>(() => reopened.ToBytes(WordFileFormat.Doc));
    }
}
