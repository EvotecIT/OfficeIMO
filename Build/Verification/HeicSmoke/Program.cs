using OfficeIMO.Drawing;
using OfficeIMO.Drawing.ImageIO;
using OfficeIMO.Drawing.ImageIO.Tests;
using OfficeIMO.Pdf;

if (args.Length == 3 && args[0] == "--compare") {
    OfficeRasterImage expected = OfficeRasterImageDecoder.Decode(File.ReadAllBytes(args[1]));
    OfficeRasterImage actual = OfficeRasterImageDecoder.Decode(File.ReadAllBytes(args[2]));
    if (actual.Width != expected.Width || actual.Height != expected.Height)
        throw new InvalidOperationException($"Dimensions differ: {actual.Width}x{actual.Height}, expected {expected.Width}x{expected.Height}.");
    byte[] a = actual.GetPixels(), b = expected.GetPixels();
    long sum = 0, over8 = 0; int maximum = 0;
    for (int i = 0; i < a.Length; i++) {
        int difference = Math.Abs(a[i] - b[i]); sum += difference;
        maximum = Math.Max(maximum, difference); if (difference > 8) over8++;
    }
    double mean = sum / (double)a.Length, fraction = over8 / (double)a.Length;
    Console.WriteLine(FormattableString.Invariant($"{actual.Width}x{actual.Height}: mean={mean:F6}, max={maximum}, fractionOver8={fraction:F6}"));
    if (mean > 2 || fraction > .01) throw new InvalidOperationException("Rendered PDF differs from the independent system reference.");
    return;
}
if (args.Length != 2) throw new ArgumentException("Usage: HeicSmoke <fixture-directory> <output-directory>, or --compare <reference.png> <rendered.png>.");
OfficeImageIOCodec.Register();
Directory.CreateDirectory(args[1]);
foreach (string name in new[] { "basic-8bit-srgb.heic", "orientation-6.heic", "display-p3-10bit.heic" }) {
    string path = Path.Combine(args[0], name), stem = Path.GetFileNameWithoutExtension(name);
    byte[] source = File.ReadAllBytes(path);
    var image = OfficeRasterImageDecoder.Decode(source);
    File.WriteAllBytes(Path.Combine(args[1], stem + ".composed.png"), OfficePngWriter.Encode(image));
    using var pdf = new[] { new PdfImageDocumentSource(source, name) }.ToPdf();
    byte[] bytes = pdf.ToBytes();
    if (!bytes.AsSpan(0, 5).SequenceEqual("%PDF-"u8)) throw new InvalidOperationException("Invalid PDF output.");
    File.WriteAllBytes(Path.Combine(args[1], stem + ".pdf"), bytes);
    if (OperatingSystem.IsMacOS()) {
        OfficeRasterImage reference = AppleImageIOReference.DecodeFullFile(source, 16_000_000);
        File.WriteAllBytes(Path.Combine(args[1], stem + ".reference.png"), OfficePngWriter.Encode(reference));
        Console.WriteLine($"{name}: {AppleImageIOReference.DescribeStoredImage(source)}");
    }
    Console.WriteLine($"{name}: {image.Width}x{image.Height}, PDF {bytes.Length} bytes");
}
