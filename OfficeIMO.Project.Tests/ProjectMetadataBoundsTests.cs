using System.Collections.Generic;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectMetadataBoundsTests {
    private static byte[] Property(ushort type, uint length) {
        var bytes = new byte[8];
        Buffer.BlockCopy(BitConverter.GetBytes(type), 0, bytes, 0, 2);
        Buffer.BlockCopy(BitConverter.GetBytes(length), 0, bytes, 4, 4);
        return bytes;
    }
    private static byte[] Set(params (uint Id, byte[] Value)[] values) {
        using var section = new MemoryStream(); using var writer = new BinaryWriter(section);
        writer.Write(8 + values.Length * 8 + values.Sum(v => v.Value.Length)); writer.Write(values.Length);
        int offset = 8 + values.Length * 8;
        foreach (var value in values) { writer.Write(value.Id); writer.Write(offset); offset += value.Value.Length; }
        foreach (var value in values) writer.Write(value.Value);
        return OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId, section.ToArray()));
    }
    [Theory]
    [InlineData(0x41, 2147483647u)]
    [InlineData(0x1e, 2147483647u)]
    [InlineData(0x1f, 2147483647u)]
    [InlineData(0x41, 4294967295u)]
    [InlineData(0x1e, 4294967295u)]
    [InlineData(0x1f, 4294967295u)]
    public void MalformedMetadataCannotAllocateBeyondItsNativeInput(ushort type, uint length) {
        byte[] metadata = Set((2, Property(type, length)));
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetReader.ReadSections(metadata));
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp")), out OfficeCompoundFile? file, out var error), error);
        byte[] mpp = OfficeCompoundFileWriter.Rewrite(file!, new Dictionary<string, byte[]> {
            [OfficeOlePropertySetWriter.SummaryInformationStreamName] = metadata
        });
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(mpp), new ProjectLoadOptions { MaxInputBytes = mpp.Length }));
    }
    [Theory]
    [InlineData(0x41)] [InlineData(0x1e)] [InlineData(0x1f)] [InlineData(5)]
    public void ValuesCannotBorrowBytesFromTheNextProperty(ushort type) {
        // The first property has only eight bytes. Its declared payload would fit the
        // whole stream, but crosses into the valid sibling string.
        var bytes = Set((2, Property(type, 4)), (3, OfficeOleProperty.String(3, "Sibling value").ValueBytes));
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetReader.ReadSections(bytes));
    }
    [Theory]
    [InlineData(44, 2147483647u)]
    [InlineData(44, 4294967295u)]
    [InlineData(48, 4294967295u)]
    [InlineData(52, 4294967295u)]
    [InlineData(60, 4u)]
    [InlineData(60, 4294967295u)]
    public void InvalidSectionAndPropertyExtentsAreRejected(int offset, uint value) {
        var bytes = Set((2, OfficeOleProperty.String(2, "Valid").ValueBytes));
        Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 4);
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetReader.ReadSections(bytes));
    }
    [Theory]
    [InlineData(1252)] [InlineData(1200)] [InlineData(65001)]
    public void DictionaryNamesAreBoundedByTheirProperty(int codePage) {
        var dictionary = new byte[12];
        Buffer.BlockCopy(BitConverter.GetBytes(1), 0, dictionary, 0, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(2), 0, dictionary, 4, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, dictionary, 8, 4);
        var bytes = Set((1, Property(2, (uint)codePage)), (0, dictionary));
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetReader.ReadSections(bytes));
    }
    [Fact]
    public void ValidMetadataStillDecodesAndCancellationPrecedesParsing() {
        var bytes = Set((2, OfficeOleProperty.String(2, "Łódź 日本語").ValueBytes));
        Assert.Equal("Łódź 日本語", OfficeOlePropertySetReader.ReadSections(bytes).Single().Properties[2].AsString());
        Assert.Throws<OperationCanceledException>(() => OfficeOlePropertySetReader.ReadSections(bytes, new CancellationToken(true)));
    }
}
