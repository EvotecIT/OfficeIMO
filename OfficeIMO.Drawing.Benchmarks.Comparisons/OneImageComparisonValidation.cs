#if NET10_0_OR_GREATER
using OneImage;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

internal static partial class ImageComparisonValidation {
    /// <summary>OneImage must decode the PNG to exactly the RGBA pixels every other engine produces.</summary>
    internal static void ValidateOneImagePngDecode(byte[] encoded) {
        byte[] expected = ImageBenchmarkCorpus.Decode(encoded, "OfficeIMO PNG validation").GetPixels();
        AssertPixels("OneImage", expected, OneImageComparisonAdapters.DecodeRgba(encoded));
    }

    /// <summary>OneImage must match the JPEG dimensions and stay within the decode error limit the other engines meet.</summary>
    internal static double ValidateOneImageJpegDecode(byte[] encoded, int width, int height) {
        byte[] actual = OneImageComparisonAdapters.DecodeRgba(encoded);
        if (actual.Length != checked(width * height * 4)) {
            throw new InvalidOperationException(
                $"OneImage did not produce the expected {width}x{height} JPEG dimensions ({actual.Length:N0} RGBA bytes).");
        }
        byte[] expected = ImageBenchmarkCorpus.Decode(encoded, "OfficeIMO JPEG validation").GetPixels();
        double error = CalculateMeanAbsoluteError(expected, actual);
        const double maximumMeanAbsoluteError = 1.5D;
        if (error > maximumMeanAbsoluteError) {
            throw new InvalidOperationException(
                $"OneImage JPEG decode exceeded the {maximumMeanAbsoluteError:F1} mean absolute channel error limit: {error:F3}.");
        }
        return error;
    }

    /// <summary>OneImage must resize to the target size within the tolerance the other resize lanes meet.</summary>
    internal static void ValidateOneImageResize(OfficeRasterImage source, byte[] expected, int width, int height) {
        using var resizer = new PixelResizer();
        AssertSimilarPixels(
            "OneImage resize",
            expected,
            OneImageComparisonAdapters.ResizeRgba(resizer, source.GetPixels(), source.Width, source.Height, width, height),
            maximumMeanAbsoluteError: 3D);
    }

    /// <summary>
    /// OneImage output must be identified and decoded exactly by OfficeIMO and Magick.NET, and OneImage must decode the
    /// OfficeIMO and Magick.NET output exactly: the same mutual-decoding contract the other lossless lanes meet.
    /// </summary>
    internal static void ValidateOneImageLosslessInterchange(
        string operation,
        OfficeRasterImage source,
        byte[] oneImageEncoded,
        byte[] officeEncoded,
        byte[] externalEncoded,
        OfficeImageFormat format) {
        byte[] expected = source.GetPixels();
        ImageBenchmarkCorpus.AssertIdentified(oneImageEncoded, format, source.Width, source.Height, operation);
        if (!OfficeRasterImageDecoder.TryDecode(oneImageEncoded, out OfficeRasterImage? decoded) || decoded == null) {
            throw new InvalidOperationException(operation + " OfficeIMO could not decode OneImage output.");
        }
        AssertPixels(operation + " OfficeIMO decode of OneImage", expected, decoded.GetPixels());
        AssertPixels(operation + " Magick.NET decode of OneImage", expected, ImageComparisonAdapters.DecodeMagickRgba(oneImageEncoded));
        AssertPixels(operation + " OneImage decode of OneImage", expected, OneImageComparisonAdapters.DecodeRgba(oneImageEncoded));
        AssertPixels(operation + " OneImage decode of OfficeIMO", expected, OneImageComparisonAdapters.DecodeRgba(officeEncoded));
        AssertPixels(operation + " OneImage decode of Magick.NET", expected, OneImageComparisonAdapters.DecodeRgba(externalEncoded));
    }

    private static void ValidateOneImage(TextWriter writer) {
        byte[] logo = ImageBenchmarkCorpus.Logo.ReadBytes();
        ValidateOneImagePngDecode(logo);
        writer.WriteLine("PNG decode: OneImage produced the identical 512x512 RGBA pixels.");

        byte[] photo = ImageBenchmarkCorpus.Photo.ReadBytes();
        double jpegError = ValidateOneImageJpegDecode(photo, 2048, 1619);
        writer.WriteLine($"JPEG decode: OneImage produced 2048x1619 RGBA; mean absolute error versus OfficeIMO was {jpegError:F3}.");

        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern();
        byte[] png = OneImageComparisonAdapters.EncodePng(
            OneImageComparisonAdapters.CreateImage(source.GetPixels(), source.Width, source.Height));
        ValidatePngEncode(source, png);
        writer.WriteLine($"PNG encode: OneImage round-tripped exact RGBA in {png.Length:N0} bytes.");

        OfficeRasterImage photoImage = ImageBenchmarkCorpus.Decode(photo, "resize validation");
        ValidateOneImageResize(photoImage, OfficeRasterResampler.Resize(photoImage, 800, 632).GetPixels(), 800, 632);
        writer.WriteLine("Resize: OneImage produced 800x632 RGBA output within the declared mean absolute error tolerance.");

        OfficeRasterImage codecSource = CreateOpaquePattern(192, 128);
        byte[] codecRgba = codecSource.GetPixels();
        global::OneImage.Image oneImage = OneImageComparisonAdapters.CreateImage(codecRgba, codecSource.Width, codecSource.Height);
        using var codecMagick = ImageComparisonAdapters.CreateMagickImage(codecRgba, codecSource.Width, codecSource.Height);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(codecMagick);
        byte[] oneTiff = OneImageComparisonAdapters.EncodeTiffLzw(oneImage);
        ValidateOneImageLosslessInterchange("TIFF LZW", codecSource, oneTiff,
            ImageComparisonAdapters.EncodeOfficeImoTiffLzw(codecSource), ImageComparisonAdapters.EncodeMagickTiffLzw(codecMagick),
            OfficeImageFormat.Tiff);
        writer.WriteLine($"TIFF LZW: OneImage, OfficeIMO and Magick.NET output decoded exactly in every engine; OneImage {oneTiff.Length:N0} bytes.");

        ImageComparisonAdapters.ConfigureMagickWebpLossless(codecMagick);
        byte[] oneWebp = OneImageComparisonAdapters.EncodeWebpLossless(oneImage);
        ValidateOneImageLosslessInterchange("lossless WebP", codecSource, oneWebp,
            ImageComparisonAdapters.EncodeOfficeImoWebpLossless(codecSource), ImageComparisonAdapters.EncodeMagickWebpLossless(codecMagick),
            OfficeImageFormat.Webp);
        writer.WriteLine($"Lossless WebP: OneImage, OfficeIMO and Magick.NET output decoded exactly in every engine; OneImage {oneWebp.Length:N0} bytes.");
    }
}
#endif
