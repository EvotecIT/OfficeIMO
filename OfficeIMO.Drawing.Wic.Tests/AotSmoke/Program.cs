using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Wic;

var codec = new OfficeWicCodec();
OfficeRasterCodecAvailability availability = codec.GetAvailability();
Console.WriteLine($"{availability.Reason}: {availability.Message}");
string? expected = args.FirstOrDefault() ?? Environment.GetEnvironmentVariable("OFFICEIMO_WIC_EXPECT_AVAILABILITY");
if (expected != null && (!Enum.TryParse(expected, out OfficeRasterCodecUnavailableReason reason) || reason != availability.Reason)) return 1;
if (!OfficeWicCodec.IsSupported) {
    IOfficeRasterImageCodec? previous = OfficeRasterImageCodecs.Platform;
    OfficeWicCodec.Register();
    return ReferenceEquals(previous, OfficeRasterImageCodecs.Platform) &&
        !codec.TryDecode(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Probe.heic")), null, out _) ? 0 : 2;
}
if (!availability.IsAvailable) return expected == null ? 3 : 0;
if (!codec.TryDecode(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Probe.heic")), "image/heic", out OfficeRasterImage? image) || image == null) return 4;
if (image.Width != 16 || image.Height != 16) return 5;
OfficeColor pixel = image.GetPixel(0, 0);
Console.WriteLine($"{image.Width}x{image.Height} RGBA {pixel.R},{pixel.G},{pixel.B},{pixel.A}");
return Math.Abs(pixel.R - 96) <= 10 && Math.Abs(pixel.G - 144) <= 10 && Math.Abs(pixel.B - 192) <= 10 && pixel.A == 255 ? 0 : 6;
