using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, true)]
    public void LegacyDoc_InlinePictureAcceptsRecoveredArrayHeadersAndRejectsTruncatedData(int count, bool truncated) {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        using WordDocument source = WordDocument.Create();
        using var imageStream = new MemoryStream(png);
        source.AddParagraph("Inline picture").AddImage(imageStream, "picture.png", 40, 20);
        source.Images[0].CropLeft = 25000;
        byte[] original = source.ToBytes(WordFileFormat.Doc);
        using WordDocument control = WordDocument.Load(new MemoryStream(original));
        Assert.Equal(png, Assert.Single(control.Images).ToBytes());
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out string? error), error);
        byte[] data = compound!.Streams["Data"];
        int shapeOffset = BitConverter.ToUInt16(data, 4);
        Assert.Equal(0xF004, BitConverter.ToUInt16(data, shapeOffset + 2));
        int shapeLength = BitConverter.ToInt32(data, shapeOffset + 4);
        int shapeEnd = shapeOffset + 8 + shapeLength;
        using var propertyStream = new MemoryStream(); using var writer = new BinaryWriter(propertyStream);
        writer.Write((ushort)0x13); writer.Write((ushort)0xF121);
        writer.Write(12 + (count == 0 ? 0 : 2));
        writer.Write((ushort)0x8146); writer.Write((uint)(count * 2));
        writer.Write((ushort)count); writer.Write((ushort)count); writer.Write((ushort)2);
        if (count > 0) writer.Write((ushort)0x8000);
        byte[] propertyRecord = propertyStream.ToArray();
        byte[] alteredData = data.Take(shapeEnd).Concat(propertyRecord).Concat(data.Skip(shapeEnd)).ToArray();
        BitConverter.GetBytes(alteredData.Length).CopyTo(alteredData, 0);
        BitConverter.GetBytes(shapeLength + propertyRecord.Length).CopyTo(alteredData, shapeOffset + 4);
        byte[] altered = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Data"] = alteredData });

        using WordDocument imported = WordDocument.Load(new MemoryStream(altered));

        if (truncated) {
            Assert.Empty(imported.Images);
        } else {
            Assert.Equal(png, Assert.Single(imported.Images).ToBytes());
            Assert.Equal(25000, imported.Images[0].CropLeft);
            Assert.Empty(imported.LegacyDocUnsupportedFeatures);
        }
    }
}
