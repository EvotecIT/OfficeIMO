using StbImageSharp;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

internal static class ImageComparisonValidation {
    internal static void Validate(TextWriter writer) {
        byte[] logo = ImageBenchmarkCorpus.Logo.ReadBytes();
        ValidatePngDecode(logo);
        writer.WriteLine("PNG decode: OfficeIMO, SkiaSharp, Magick.NET, and StbImageSharp produced identical 512x512 RGBA pixels.");

        byte[] photo = ImageBenchmarkCorpus.Photo.ReadBytes();
        (double skiaError, double magickError, double stbError) = MeasureJpegDecodeError(photo, 2048, 1619);
        writer.WriteLine(
            $"JPEG decode: expected 2048x1619 dimensions; RGBA mean absolute error versus OfficeIMO was " +
            $"SkiaSharp {skiaError:F3}, Magick.NET {magickError:F3}, StbImageSharp {stbError:F3}.");

        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern();
        byte[] rgba = source.GetPixels();
        using var skia = ImageComparisonAdapters.CreateSkiaBitmap(rgba, source.Width, source.Height);
        using var magick = ImageComparisonAdapters.CreateMagickImage(rgba, source.Width, source.Height);
        byte[] officePng = ImageComparisonAdapters.EncodeOfficeImo(source);
        byte[] skiaPng = ImageComparisonAdapters.EncodeSkia(skia);
        byte[] magickPng = ImageComparisonAdapters.EncodeMagick(magick);
        ValidatePngEncode(source, officePng, skiaPng, magickPng);
        writer.WriteLine($"PNG encode: exact RGBA round-trip; OfficeIMO {officePng.Length:N0}, SkiaSharp {skiaPng.Length:N0}, Magick.NET {magickPng.Length:N0} bytes.");

        OfficeRasterImage photoImage = ImageBenchmarkCorpus.Decode(photo, "resize validation");
        byte[] photoRgba = photoImage.GetPixels();
        using var photoSkia = ImageComparisonAdapters.CreateSkiaBitmap(photoRgba, 2048, 1619);
        using var photoMagick = ImageComparisonAdapters.CreateMagickImage(photoRgba, 2048, 1619);
        byte[] officeResize = OfficeRasterResampler.Resize(photoImage, 800, 632).GetPixels();
        AssertSimilarPixels("SkiaSharp resize", officeResize, ImageComparisonAdapters.ResizeSkiaRgba(photoSkia, 800, 632), 3D);
        AssertSimilarPixels("Magick.NET resize", officeResize, ImageComparisonAdapters.ResizeMagickRgba(photoMagick, 800, 632), 3D);
        writer.WriteLine("Resize: all engines produced 800x632 RGBA output within the declared mean absolute error tolerance.");
    }

    internal static void ValidatePngDecode(byte[] encoded) {
        OfficeRasterImage office = ImageBenchmarkCorpus.Decode(encoded, "OfficeIMO PNG validation");
        byte[] expected = office.GetPixels();
        AssertPixels("SkiaSharp", expected, ImageComparisonAdapters.DecodeSkiaRgba(encoded));
        AssertPixels("Magick.NET", expected, ImageComparisonAdapters.DecodeMagickRgba(encoded));
        ImageResult stb = ImageResult.FromMemory(encoded, ColorComponents.RedGreenBlueAlpha);
        AssertPixels("StbImageSharp", expected, stb.Data);
    }

    internal static void ValidateDimensions(byte[] encoded, int width, int height) {
        int expected = checked(width * height * 4);
        if (ImageComparisonAdapters.DecodeOfficeImoRgba(encoded).Length != expected ||
            ImageComparisonAdapters.DecodeSkiaRgba(encoded).Length != expected ||
            ImageComparisonAdapters.DecodeMagickRgba(encoded).Length != expected ||
            ImageComparisonAdapters.DecodeStbRgba(encoded).Length != expected) {
            throw new InvalidOperationException("A JPEG decoder did not produce the expected dimensions.");
        }
    }

    internal static (double SkiaError, double MagickError, double StbError) MeasureJpegDecodeError(
        byte[] encoded,
        int width,
        int height) {
        ValidateDimensions(encoded, width, height);
        OfficeRasterImage office = ImageBenchmarkCorpus.Decode(encoded, "OfficeIMO JPEG validation");
        byte[] expected = office.GetPixels();
        ImageResult stb = ImageResult.FromMemory(encoded, ColorComponents.RedGreenBlueAlpha);
        var errors = (
            SkiaError: CalculateMeanAbsoluteError(expected, ImageComparisonAdapters.DecodeSkiaRgba(encoded)),
            MagickError: CalculateMeanAbsoluteError(expected, ImageComparisonAdapters.DecodeMagickRgba(encoded)),
            StbError: CalculateMeanAbsoluteError(expected, stb.Data));
        const double maximumMeanAbsoluteError = 1.5D;
        if (errors.SkiaError > maximumMeanAbsoluteError ||
            errors.MagickError > maximumMeanAbsoluteError ||
            errors.StbError > maximumMeanAbsoluteError) {
            throw new InvalidOperationException(
                $"JPEG decoder output exceeded the {maximumMeanAbsoluteError:F1} mean absolute channel error limit: " +
                $"SkiaSharp {errors.SkiaError:F3}, Magick.NET {errors.MagickError:F3}, StbImageSharp {errors.StbError:F3}.");
        }
        return errors;
    }

    internal static void ValidatePngEncode(OfficeRasterImage source, params byte[][] encodedImages) {
        byte[] expected = source.GetPixels();
        foreach (byte[] encoded in encodedImages) {
            ImageBenchmarkCorpus.AssertIdentified(encoded, OfficeImageFormat.Png, source.Width, source.Height, "PNG comparison encode");
            OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(encoded, "PNG comparison encode");
            AssertPixels("PNG encoder", expected, decoded.GetPixels());
        }
    }

    private static void AssertPixels(string engine, byte[] expected, byte[] actual) {
        if (!expected.AsSpan().SequenceEqual(actual)) {
            throw new InvalidOperationException(engine + " did not produce the expected RGBA pixels.");
        }
    }

    internal static void AssertSimilarPixels(string engine, byte[] expected, byte[] actual, double maximumMeanAbsoluteError) {
        double meanAbsoluteError = CalculateMeanAbsoluteError(expected, actual);
        if (meanAbsoluteError > maximumMeanAbsoluteError) {
            throw new InvalidOperationException(
                $"{engine} mean absolute channel error {meanAbsoluteError:F3} exceeded {maximumMeanAbsoluteError:F3}.");
        }
    }

    private static double CalculateMeanAbsoluteError(byte[] expected, byte[] actual) {
        if (expected.Length != actual.Length) {
            throw new InvalidOperationException("The image decoder did not produce the expected pixel count.");
        }
        long absoluteError = 0L;
        for (int index = 0; index < expected.Length; index++) {
            absoluteError += Math.Abs(expected[index] - actual[index]);
        }
        return absoluteError / (double)expected.Length;
    }
}
