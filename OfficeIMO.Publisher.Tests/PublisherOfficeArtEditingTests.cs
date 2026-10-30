using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherOfficeArtEditingTests {
    [Fact]
    public void Scalar_table_growth_preserves_unedited_rows_and_the_raw_complex_tail() {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)0x81); writer.Write(0U);
        writer.Write((ushort)0x83FE); writer.Write(4U);
        writer.Write((ushort)0x3FF); writer.Write(42U);
        byte[] tail = { 1, 2, 3, 4, 0xCC }; writer.Write(tail);
        byte[] source = stream.ToArray();
        byte[]? edited = OfficeArtPropertyTableEditor.RewriteScalars(source, 0, source.Length, 3,
            new Dictionary<ushort, uint> { [0x81] = 152400, [0x8C] = 2, [0x8D] = 152400 }, out ushort count);
        Assert.NotNull(edited); Assert.Equal(5, count);
        var properties = OfficeArtPropertyTableReader.Read(edited!, count);
        Assert.Equal(new ushort[] { 0x81, 0x8C, 0x8D, 0x3FE, 0x3FF }, properties.Select(property => property.PropertyId));
        Assert.Equal(new uint[] { 152400, 2, 152400, 4, 42 }, properties.Select(property => property.Value));
        Assert.Equal(tail, edited!.Skip(count * 6));
        Assert.Equal(source.Skip(6).Take(12), edited.Skip(18).Take(12));
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("complex-target")]
    [InlineData("truncated-tail")]
    public void Scalar_table_editing_refuses_ambiguous_targets_or_incomplete_complex_data(string invalid) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)(invalid == "complex-target" ? 0x8081 : 0x81)); writer.Write(0U);
        writer.Write((ushort)(invalid == "duplicate" ? 0x81 : 0x83FE)); writer.Write(invalid == "duplicate" ? 0U : 4U);
        if (invalid != "truncated-tail") writer.Write(new byte[] { 1, 2, 3, 4 });
        byte[] source = stream.ToArray();
        Assert.Null(OfficeArtPropertyTableEditor.RewriteScalars(source, 0, source.Length, 2,
            new Dictionary<ushort, uint> { [0x81] = 12700 }, out _));
    }
}
