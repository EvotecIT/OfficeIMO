using System.Security.Cryptography;

namespace OfficeIMO.Drawing.Benchmarks;

internal sealed record ImageBenchmarkAsset(
    string Id,
    string FileName,
    OfficeImageFormat Format,
    int Width,
    int Height) {
    internal byte[] ReadBytes() => File.ReadAllBytes(ImageBenchmarkCorpus.GetPath(FileName));
}

internal static class ImageBenchmarkCorpus {
    internal static readonly ImageBenchmarkAsset Logo = new("LogoPng", "logo.png", OfficeImageFormat.Png, 512, 512);
    internal static readonly ImageBenchmarkAsset Photo = new("PhotoJpeg", "photo.jpg", OfficeImageFormat.Jpeg, 2048, 1619);
    internal static readonly ImageBenchmarkAsset Animation = new("AnimationGif", "animation.gif", OfficeImageFormat.Gif, 220, 137);
    internal static readonly ImageBenchmarkAsset Tiff = new("SaturnTiff", "saturn.tif", OfficeImageFormat.Tiff, 640, 480);
    internal static readonly ImageBenchmarkAsset Bitmap = new("SnailBmp", "snail.bmp", OfficeImageFormat.Bmp, 256, 256);

    internal static IReadOnlyList<ImageBenchmarkAsset> All { get; } = [Logo, Photo, Animation, Tiff, Bitmap];

    internal static ImageBenchmarkAsset Get(string id) => All.Single(asset => asset.Id == id);

    internal static string GetPath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Corpus", fileName);

    internal static OfficeRasterImage CreatePattern(int width = 512, int height = 512) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                byte red = (byte)((x * 13 + y * 3) & 255);
                byte green = (byte)((x * 5 + y * 11) & 255);
                byte blue = (byte)(((x ^ y) * 7) & 255);
                byte alpha = (byte)(96 + ((x + y) & 159));
                image.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, alpha));
            }
        }
        return image;
    }

    internal static byte[] CreateBmp24(int width = 256, int height = 256) {
        int rowStride = checked(((width * 24) + 31) / 32 * 4);
        int pixelOffset = 54;
        var bytes = new byte[checked(pixelOffset + (rowStride * height))];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        WriteInt32LittleEndian(bytes, 2, bytes.Length);
        WriteInt32LittleEndian(bytes, 10, pixelOffset);
        WriteInt32LittleEndian(bytes, 14, 40);
        WriteInt32LittleEndian(bytes, 18, width);
        WriteInt32LittleEndian(bytes, 22, height);
        bytes[26] = 1;
        bytes[28] = 24;
        WriteInt32LittleEndian(bytes, 34, rowStride * height);
        for (int y = 0; y < height; y++) {
            int row = pixelOffset + ((height - 1 - y) * rowStride);
            for (int x = 0; x < width; x++) {
                int pixel = row + (x * 3);
                bytes[pixel] = (byte)(((x ^ y) * 7) & 255);
                bytes[pixel + 1] = (byte)((x * 5 + y * 11) & 255);
                bytes[pixel + 2] = (byte)((x * 13 + y * 3) & 255);
            }
        }
        return bytes;
    }

    internal static string PixelHash(OfficeRasterImage image) =>
        Convert.ToHexString(SHA256.HashData(image.GetPixels()));

    internal static void AssertIdentified(byte[] encoded, OfficeImageFormat format, int width, int height, string operation) {
        if (!OfficeImageReader.TryIdentify(encoded, fileName: null, out OfficeImageInfo info)) {
            throw new InvalidOperationException(operation + " did not produce identifiable image bytes.");
        }
        if (info.Format != format || info.Width != width || info.Height != height) {
            throw new InvalidOperationException(
                $"{operation} produced {info.Format} {info.Width}x{info.Height}; expected {format} {width}x{height}.");
        }
    }

    internal static OfficeRasterImage Decode(byte[] encoded, string operation) {
        if (!OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? image) || image == null) {
            throw new InvalidOperationException(operation + " could not be decoded by the managed image engine.");
        }
        return image;
    }

    private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }
}
