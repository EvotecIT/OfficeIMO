using System.Text;
using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(0x0145, 8)]
    [InlineData(0x0145, 4)]
    [InlineData(0x0145, 0xFFF0)]
    [InlineData(0x0146, 2)]
    [InlineData(0x0156, 8)]
    [InlineData(0x0197, 8)]
    public void OfficeArtPropertyTableReader_ArrayLengthExcludingHeaderPreservesTheFollowingProperty(ushort id, int elementSize) {
        int stride = elementSize == 0xFFF0 ? 4 : elementSize;
        byte[] array = new byte[6 + 2 * stride];
        array[0] = array[2] = 2;
        array[4] = (byte)elementSize; array[5] = (byte)(elementSize >> 8);
        byte[] name = Encoding.Unicode.GetBytes("After");
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)(id | 0x8000)); writer.Write((uint)(array.Length - 6));
        writer.Write((ushort)0x8380); writer.Write((uint)name.Length);
        writer.Write(array); writer.Write(name);

        var properties = OfficeArtPropertyTableReader.Read(stream.ToArray(), 2);

        Assert.Equal((uint)(array.Length - 6), properties[0].DeclaredComplexDataLength);
        Assert.Equal(array.Length, properties[0].AvailableComplexDataLength);
        Assert.Equal(array, properties[0].CopyComplexData());
        Assert.True(properties[0].HasCompleteComplexData);
        Assert.Equal("After", properties[1].ComplexText);
    }

    [Theory]
    [InlineData(0x0145, 8)]
    [InlineData(0x0145, 4)]
    [InlineData(0x0145, 0xFFF0)]
    [InlineData(0x0146, 2)]
    [InlineData(0x0156, 8)]
    [InlineData(0x0197, 8)]
    public void OfficeArtPropertyTableReader_EmptyArrayHeaderPreservesTheFollowingProperty(ushort id, ushort elementSize) {
        byte[] name = Encoding.Unicode.GetBytes("After");
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)(id | 0x8000)); writer.Write(0U);
        writer.Write((ushort)0x8380); writer.Write((uint)name.Length);
        writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(elementSize);
        writer.Write(name);

        var properties = OfficeArtPropertyTableReader.Read(stream.ToArray(), 2);

        Assert.Equal(0U, properties[0].DeclaredComplexDataLength);
        Assert.Equal(6, properties[0].AvailableComplexDataLength);
        Assert.True(properties[0].HasCompleteComplexData);
        Assert.Equal("After", properties[1].ComplexText);
        Assert.True(properties[1].HasCompleteComplexData);
    }

    [Fact]
    public void OfficeArtPropertyTableReader_AbsentArrayDoesNotConsumeTheNextDeclaredPayload() {
        byte[] nextData = { 0, 0, 0, 0, 2, 0 };
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)0x8146); writer.Write(0U);
        writer.Write((ushort)0x83FE); writer.Write((uint)nextData.Length);
        writer.Write(nextData);

        var properties = OfficeArtPropertyTableReader.Read(stream.ToArray(), 2);

        Assert.Equal(0, properties[0].AvailableComplexDataLength);
        Assert.Null(properties[0].CopyComplexData());
        Assert.Equal(nextData, properties[1].CopyComplexData());
        Assert.True(properties[1].HasCompleteComplexData);
    }

    [Fact]
    public void OfficeArtPropertyTableReader_DoesNotGuessArrayHeadersForUnrelatedComplexData() {
        byte[] payload = { 0xFE, 0x83, 8, 0, 0, 0, 2, 0, 2, 0, 4, 0, 7, 0, 9, 0, 0, 0, 0, 0 };
        OfficeArtProperty property = Assert.Single(OfficeArtPropertyTableReader.Read(payload, 1));
        Assert.Equal(8, property.AvailableComplexDataLength);
        Assert.Equal(payload.Skip(6).Take(8), property.CopyComplexData());
    }

    [Fact]
    public void OfficeArtPropertyTableReader_TruncatedArrayCannotReadPastTheContainingRecord() {
        byte[] payload = { 0x45, 0x81, 8, 0, 0, 0, 1, 0, 1, 0, 8, 0, 7, 0, 9, 0 };
        OfficeArtProperty property = Assert.Single(OfficeArtPropertyTableReader.Read(payload, 1));
        Assert.Equal(10, property.AvailableComplexDataLength);
        Assert.Equal(payload.Skip(6), property.CopyComplexData());
        Assert.False(property.HasCompleteComplexData);
    }
}
