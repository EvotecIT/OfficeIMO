using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffArithmeticTableTests {
    [Theory]
    [InlineData(8, 204)]
    [InlineData(12, 204)]
    [InlineData(16, 204)]
    [InlineData(8, 221)]
    [InlineData(12, 221)]
    [InlineData(16, 221)]
    public void TableOnlyArithmeticControlsResetBeforeTheImage(int bits, byte marker) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets",
            "JpegArithmeticLossless", $"b{bits}-c1-r0.jpg"));
        // Rely on JPEG's default conditioning and restart state. Local definitions
        // would mask a regression that incorrectly inherited table-only controls.
        for (int at = 2; jpeg[at + 1] != 218; at += 2 + (jpeg[at + 2] << 8 | jpeg[at + 3])) {
            if (jpeg[at + 1] is 204 or 221) jpeg[at + 1] = 254;
        }
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var expected));
        byte[] control = { 255, marker, 0, 4, 0, marker == 204 ? (byte)255 : (byte)1 };
        byte[] tables = new byte[] { 255, 216 }.Concat(control).Concat(new byte[] { 255, 217 }).ToArray();
        byte[] tiff = Wrap(jpeg, bits, tables);
        Assert.True(OfficeImageReader.TryValidateContent(tiff, "tables.tif", out _));
        Assert.True(OfficeTiffCodec.TryDecode(tiff, out var actual));
        for (int y = 0; y < 11; y++) for (int x = 0; x < 19; x++)
            Assert.Equal(expected!.GetPixel(x, y), actual!.GetPixel(x, y));

        // Confirm this stream actually distinguishes reset from inherited state.
        byte[] inherited = jpeg.Take(2).Concat(control).Concat(jpeg.Skip(2)).ToArray();
        bool decoded = OfficeJpegCodec.TryDecode(inherited, out var poisoned);
        Assert.True(!decoded || Enumerable.Range(0, 19 * 11).Any(i =>
            poisoned!.GetPixel(i % 19, i / 19) != expected!.GetPixel(i % 19, i / 19)));
    }

    private static byte[] Wrap(byte[] jpeg, int bits, byte[] tables) {
        const int count = 11, tableOffset = 8 + 2 + count * 12 + 4;
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output);
        writer.Write((ushort)0x4949); writer.Write((ushort)42); writer.Write(8U); writer.Write((ushort)count);
        void Entry(ushort tag, ushort type, uint length, uint value) {
            writer.Write(tag); writer.Write(type); writer.Write(length); writer.Write(value);
        }
        Entry(256, 4, 1, 19); Entry(257, 4, 1, 11); Entry(258, 3, 1, (uint)bits);
        Entry(259, 3, 1, 7); Entry(262, 3, 1, 1); Entry(273, 4, 1, (uint)(tableOffset + tables.Length));
        Entry(277, 3, 1, 1); Entry(278, 4, 1, 11); Entry(279, 4, 1, (uint)jpeg.Length);
        Entry(284, 3, 1, 1); Entry(347, 7, (uint)tables.Length, tableOffset);
        writer.Write(0U); writer.Write(tables); writer.Write(jpeg);
        return output.ToArray();
    }
}
