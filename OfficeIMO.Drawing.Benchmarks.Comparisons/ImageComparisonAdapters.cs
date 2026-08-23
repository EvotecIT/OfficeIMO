using ImageMagick;
using SkiaSharp;
using StbImageSharp;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

internal static class ImageComparisonAdapters {
    internal static byte[] DecodeOfficeImoRgba(byte[] encoded) {
        OfficeRasterImage image = ImageBenchmarkCorpus.Decode(encoded, "OfficeIMO decode");
        return image.GetPixels();
    }

    internal static byte[] DecodeSkiaRgba(byte[] encoded) {
        using SKBitmap decoded = DecodeSkiaBitmap(encoded);
        var rgba = new byte[checked(decoded.Width * decoded.Height * 4)];
        System.Runtime.InteropServices.Marshal.Copy(decoded.GetPixels(), rgba, 0, rgba.Length);
        return rgba;
    }

    internal static byte[] DecodeMagickRgba(byte[] encoded) {
        using var image = new MagickImage(encoded);
        using IPixelCollection<byte> pixels = image.GetPixels();
        return pixels.ToByteArray(PixelMapping.RGBA)
            ?? throw new InvalidOperationException("Magick.NET did not expose decoded RGBA pixels.");
    }

    internal static byte[] DecodeStbRgba(byte[] encoded) {
        ImageResult image = ImageResult.FromMemory(encoded, ColorComponents.RedGreenBlueAlpha);
        return image.Data;
    }

    internal static byte[] EncodeOfficeImo(OfficeRasterImage image) =>
        OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png);

    internal static byte[] EncodeSkia(SKBitmap bitmap) {
        using SKImage image = SKImage.FromBitmap(bitmap);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    internal static byte[] EncodeMagick(MagickImage image) {
        image.Format = MagickFormat.Png;
        return image.ToByteArray();
    }

    internal static SKBitmap CreateSkiaBitmap(byte[] rgba, int width, int height) {
        var bitmap = new SKBitmap(new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Unpremul));
        System.Runtime.InteropServices.Marshal.Copy(rgba, 0, bitmap.GetPixels(), rgba.Length);
        return bitmap;
    }

    internal static MagickImage CreateMagickImage(byte[] rgba, int width, int height) =>
        new(rgba, new PixelReadSettings((uint)width, (uint)height, StorageType.Char, PixelMapping.RGBA));

    internal static int ResizeOfficeImo(OfficeRasterImage image, int width, int height) {
        OfficeRasterImage resized = OfficeRasterResampler.Resize(image, width, height, OfficeRasterResamplingMode.Bilinear);
        return checked(resized.Width * resized.Height);
    }

    internal static int ResizeSkia(SKBitmap image, int width, int height) {
        using SKBitmap resized = image.Resize(
            new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Unpremul),
            new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.None))
            ?? throw new InvalidOperationException("SkiaSharp could not resize the image.");
        return checked(resized.Width * resized.Height);
    }

    internal static int ResizeMagick(MagickImage image, int width, int height) {
        using IMagickImage<byte> resized = image.Clone();
        resized.FilterType = FilterType.Triangle;
        resized.Resize(new MagickGeometry((uint)width, (uint)height) { IgnoreAspectRatio = true });
        return checked((int)(resized.Width * resized.Height));
    }

    internal static byte[] ResizeSkiaRgba(SKBitmap image, int width, int height) {
        using SKBitmap resized = image.Resize(
            new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Unpremul),
            new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.None))
            ?? throw new InvalidOperationException("SkiaSharp could not resize the image.");
        var rgba = new byte[checked(width * height * 4)];
        System.Runtime.InteropServices.Marshal.Copy(resized.GetPixels(), rgba, 0, rgba.Length);
        return rgba;
    }

    internal static byte[] ResizeMagickRgba(MagickImage image, int width, int height) {
        using IMagickImage<byte> resized = image.Clone();
        resized.FilterType = FilterType.Triangle;
        resized.Resize(new MagickGeometry((uint)width, (uint)height) { IgnoreAspectRatio = true });
        return resized.ToByteArray(MagickFormat.Rgba);
    }

    private static SKBitmap DecodeSkiaBitmap(byte[] encoded) {
        using SKData data = SKData.CreateCopy(encoded);
        using SKCodec codec = SKCodec.Create(data)
            ?? throw new InvalidOperationException("SkiaSharp could not create a decoder.");
        var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Unpremul);
        var decoded = new SKBitmap(info);
        SKCodecResult result = codec.GetPixels(info, decoded.GetPixels());
        if (result == SKCodecResult.Success) return decoded;
        decoded.Dispose();
        throw new InvalidOperationException("SkiaSharp did not decode a complete RGBA image: " + result + ".");
    }

}
