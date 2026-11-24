using System.Globalization;
using OfficeIMO.Benchmarks;
using StbImageSharp;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>
/// Checks every comparison lane's real output before timing and records the result for the evidence catalog.
/// Each lane is compared with a decoder or engine other than itself, so no library validates its own output.
/// </summary>
internal static partial class ImageComparisonValidation {
    internal const string OfficeImo = "OfficeIMO";
    internal const string Skia = "SkiaSharp";
    internal const string Magick = "Magick.NET";
    internal const string Stb = "StbImageSharp";
    internal const string OneImageLibrary = "OneImage";

    /// <summary>Largest full-buffer mean absolute RGBA channel error accepted between two JPEG decoders.</summary>
    internal const double JpegMaximumMeanAbsoluteError = 1.5D;

    /// <summary>Largest full-buffer mean absolute RGBA channel error accepted between two linear resizers.</summary>
    internal const double ResizeMaximumMeanAbsoluteError = 3D;

    internal static void Validate(TextWriter writer) {
        byte[] logo = ImageBenchmarkCorpus.Logo.ReadBytes();
        ValidatePngDecode(logo);
        writer.WriteLine("PNG decode: every decoder produced 512x512 RGBA pixels identical to an independent decoder.");

        byte[] photo = ImageBenchmarkCorpus.Photo.ReadBytes();
        writer.WriteLine("JPEG decode: 2048x1619 RGBA mean absolute error per decoder against the other decoders: " +
            string.Join("; ", ValidateJpegDecode(photo, 2048, 1619).Select(pair => pair.Key + " " + pair.Value)) + ".");

        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern();
        byte[] rgba = source.GetPixels();
        using var skia = ImageComparisonAdapters.CreateSkiaBitmap(rgba, source.Width, source.Height);
        using var magick = ImageComparisonAdapters.CreateMagickImage(rgba, source.Width, source.Height);
        byte[] officePng = ImageComparisonAdapters.EncodeOfficeImo(source);
        byte[] skiaPng = ImageComparisonAdapters.EncodeSkia(skia);
        byte[] magickPng = ImageComparisonAdapters.EncodeMagick(magick);
        ValidatePngEncode(source, officePng, skiaPng, magickPng);
        writer.WriteLine($"PNG encode: exact RGBA round-trip through StbImageSharp; OfficeIMO {officePng.Length:N0}, SkiaSharp {skiaPng.Length:N0}, Magick.NET {magickPng.Length:N0} bytes.");

        OfficeRasterImage photoImage = ImageBenchmarkCorpus.Decode(photo, "resize validation");
        byte[] photoRgba = photoImage.GetPixels();
        using var photoSkia = ImageComparisonAdapters.CreateSkiaBitmap(photoRgba, 2048, 1619);
        using var photoMagick = ImageComparisonAdapters.CreateMagickImage(photoRgba, 2048, 1619);
        ValidateResize(photoImage, photoSkia, photoMagick, 800, 632);
        writer.WriteLine("Resize: all engines produced 800x632 RGBA output within the declared mean absolute error tolerance of each other.");

        OfficeRasterImage codecSource = CreateOpaquePattern(192, 128);
        byte[] codecRgba = codecSource.GetPixels();
        using var codecMagick = ImageComparisonAdapters.CreateMagickImage(codecRgba, codecSource.Width, codecSource.Height);
        byte[] officeTiff = ImageComparisonAdapters.EncodeOfficeImoTiffLzw(codecSource);
        ImageComparisonAdapters.ConfigureMagickTiffLzw(codecMagick);
        byte[] magickTiff = ImageComparisonAdapters.EncodeMagickTiffLzw(codecMagick);
        ValidateLosslessEncode("TIFF LZW", codecSource, officeTiff, magickTiff, OfficeImageFormat.Tiff);
        ValidateLosslessDecode("TIFF LZW", codecSource, magickTiff, OfficeImageFormat.Tiff);
        writer.WriteLine($"TIFF LZW: OfficeIMO and Magick.NET encoded mutually decodable exact RGBA output; OfficeIMO {officeTiff.Length:N0}, Magick.NET {magickTiff.Length:N0} bytes.");

        var firstPage = new OfficeRasterImage(16, 12, OfficeColor.Crimson);
        var secondPage = new OfficeRasterImage(9, 7, OfficeColor.CornflowerBlue);
        byte[] externalPages = ImageComparisonAdapters.EncodeMagickMultiPageTiffLzw(firstPage, secondPage);
        var pageOptions = new OfficeRasterDecodeOptions { FrameIndex = 1 };
        if (!OfficeRasterImageDecoder.TryDecode(externalPages, pageOptions, out OfficeRasterImage? selectedPage, out OfficeRasterDecodeInfo pageInfo) ||
            selectedPage == null || pageInfo.FrameCount != 2 || !pageInfo.PagesDiscarded) {
            throw new InvalidOperationException("OfficeIMO did not select the second page from an external multi-page TIFF.");
        }
        AssertPixels("external multi-page TIFF selection", secondPage.GetPixels(), selectedPage.GetPixels());
        writer.WriteLine("TIFF pages: OfficeIMO inspected and selected the second LZW page from a Magick.NET/libtiff multi-page container exactly.");

        byte[] officeWebp = ImageComparisonAdapters.EncodeOfficeImoWebpLossless(codecSource);
        ImageComparisonAdapters.ConfigureMagickWebpLossless(codecMagick);
        byte[] magickWebp = ImageComparisonAdapters.EncodeMagickWebpLossless(codecMagick);
        ValidateLosslessEncode("lossless WebP", codecSource, officeWebp, magickWebp, OfficeImageFormat.Webp);
        ValidateLosslessDecode("lossless WebP", codecSource, magickWebp, OfficeImageFormat.Webp);
        writer.WriteLine($"Lossless WebP: OfficeIMO and Magick.NET encoded mutually decodable exact RGBA output; OfficeIMO {officeWebp.Length:N0}, Magick.NET {magickWebp.Length:N0} bytes.");
    }

    /// <summary>
    /// Every PNG decoder must produce the asset's dimensions and exactly the RGBA pixels of another decoder:
    /// OfficeIMO, SkiaSharp and StbImageSharp against Magick.NET, and Magick.NET against StbImageSharp.
    /// </summary>
    internal static void ValidatePngDecode(byte[] encoded) {
        ImageBenchmarkAsset asset = ImageBenchmarkCorpus.Logo;
        var decoded = new Dictionary<string, byte[]> {
            [OfficeImo] = ImageComparisonAdapters.DecodeOfficeImoRgba(encoded),
            [Skia] = ImageComparisonAdapters.DecodeSkiaRgba(encoded),
            [Magick] = ImageComparisonAdapters.DecodeMagickRgba(encoded),
            [Stb] = ImageComparisonAdapters.DecodeStbRgba(encoded),
        };
#if NET10_0_OR_GREATER
        decoded[OneImageLibrary] = OneImageComparisonAdapters.DecodeRgba(encoded);
#endif
        int expectedLength = checked(asset.Width * asset.Height * 4);
        foreach (KeyValuePair<string, byte[]> lane in decoded) {
            string reference = lane.Key == Magick ? Stb : Magick;
            Record(lane.Key, "Default", () => {
                if (lane.Value.Length != expectedLength) {
                    throw new InvalidOperationException(
                        $"{lane.Key} PNG decode produced {lane.Value.Length} RGBA bytes; expected {asset.Width}x{asset.Height}.");
                }
                AssertPixels(lane.Key + " PNG decode against " + reference, decoded[reference], lane.Value);
                return $"PNG {asset.Width}x{asset.Height}; RGBA identical to {reference}";
            });
        }
    }

    /// <summary>
    /// Every JPEG decoder must produce the expected dimensions and stay within
    /// <see cref="JpegMaximumMeanAbsoluteError"/> of every other decoder, so normal IDCT, upsampling and rounding
    /// differences pass while a material defect in any one decoder fails.
    /// </summary>
    internal static IReadOnlyDictionary<string, string> ValidateJpegDecode(byte[] encoded, int width, int height) {
        var decoded = new Dictionary<string, byte[]> {
            [OfficeImo] = ImageComparisonAdapters.DecodeOfficeImoRgba(encoded),
            [Skia] = ImageComparisonAdapters.DecodeSkiaRgba(encoded),
            [Magick] = ImageComparisonAdapters.DecodeMagickRgba(encoded),
            [Stb] = ImageComparisonAdapters.DecodeStbRgba(encoded),
        };
#if NET10_0_OR_GREATER
        decoded[OneImageLibrary] = OneImageComparisonAdapters.DecodeRgba(encoded);
#endif
        var summary = new Dictionary<string, string>();
        int expectedLength = checked(width * height * 4);
        foreach (KeyValuePair<string, byte[]> lane in decoded) {
            summary[lane.Key] = Record(lane.Key, "Default", () => {
                if (lane.Value.Length != expectedLength) {
                    throw new InvalidOperationException(
                        $"{lane.Key} JPEG decode produced {lane.Value.Length} RGBA bytes; expected {width}x{height}.");
                }
                var errors = new List<string>();
                foreach (KeyValuePair<string, byte[]> other in decoded) {
                    if (other.Key == lane.Key) continue;
                    double error = CalculateMeanAbsoluteError(other.Value, lane.Value);
                    if (error > JpegMaximumMeanAbsoluteError) {
                        throw new InvalidOperationException(
                            $"{lane.Key} JPEG decode differs from {other.Key} by mean absolute channel error {Format(error, 3)}; limit {Format(JpegMaximumMeanAbsoluteError, 1)}.");
                    }
                    errors.Add($"{other.Key} {Format(error, 3)}");
                }
                return $"JPEG {width}x{height}; RGBA mean absolute error (limit {Format(JpegMaximumMeanAbsoluteError, 1)}) versus " + string.Join(", ", errors);
            });
        }
        return summary;
    }

    /// <summary>
    /// Every PNG encoder's output must identify as a PNG of the source dimensions and decode, through StbImageSharp
    /// (which none of the encoders use), to exactly the source RGBA pixels. Outputs are OfficeIMO, SkiaSharp, Magick.NET
    /// and, where it runs, OneImage.
    /// </summary>
    internal static void ValidatePngEncode(OfficeRasterImage source, byte[] officeImo, byte[] skia, byte[] magick, byte[]? oneImage = null) {
        byte[] expected = source.GetPixels();
        var lanes = new List<(string Library, byte[] Encoded)> { (OfficeImo, officeImo), (Skia, skia), (Magick, magick) };
        if (oneImage != null) lanes.Add((OneImageLibrary, oneImage));
        foreach ((string library, byte[] encoded) in lanes) {
            Record(library, "Default", () => {
                ImageBenchmarkCorpus.AssertIdentified(encoded, OfficeImageFormat.Png, source.Width, source.Height, library + " PNG encode");
                ImageResult decoded = ImageResult.FromMemory(encoded, ColorComponents.RedGreenBlueAlpha);
                if (decoded.Width != source.Width || decoded.Height != source.Height) {
                    throw new InvalidOperationException($"StbImageSharp decoded {library} PNG output as {decoded.Width}x{decoded.Height}.");
                }
                AssertPixels(library + " PNG encode decoded by StbImageSharp", expected, decoded.Data);
                return $"PNG {source.Width}x{source.Height}, {encoded.Length} bytes; StbImageSharp decode equals source RGBA";
            });
        }
    }

    /// <summary>
    /// Every resizer must produce the target dimensions and stay within <see cref="ResizeMaximumMeanAbsoluteError"/>
    /// of every other engine's linear/bilinear output.
    /// </summary>
    internal static void ValidateResize(OfficeRasterImage officeImage, SkiaSharp.SKBitmap skiaImage, ImageMagick.MagickImage magickImage, int width, int height) {
        OfficeRasterImage officeResized = OfficeRasterResampler.Resize(officeImage, width, height, OfficeRasterResamplingMode.Bilinear);
        if (officeResized.Width != width || officeResized.Height != height) {
            BenchmarkOutputValidationLog.Failed(OfficeImo, $"Resize produced {officeResized.Width}x{officeResized.Height}; expected {width}x{height}", "Default");
            throw new InvalidOperationException($"OfficeIMO resize produced {officeResized.Width}x{officeResized.Height}; expected {width}x{height}.");
        }
        var resized = new Dictionary<string, byte[]> {
            [OfficeImo] = officeResized.GetPixels(),
            [Skia] = ImageComparisonAdapters.ResizeSkiaRgba(skiaImage, width, height),
            [Magick] = ImageComparisonAdapters.ResizeMagickRgba(magickImage, width, height),
        };
#if NET10_0_OR_GREATER
        using (var oneImageResizer = new global::OneImage.PixelResizer()) {
            resized[OneImageLibrary] = OneImageComparisonAdapters.ResizeRgba(
                oneImageResizer, officeImage.GetPixels(), officeImage.Width, officeImage.Height, width, height);
        }
#endif
        int expectedLength = checked(width * height * 4);
        foreach (KeyValuePair<string, byte[]> lane in resized) {
            Record(lane.Key, "Default", () => {
                if (lane.Value.Length != expectedLength) {
                    throw new InvalidOperationException($"{lane.Key} resize produced {lane.Value.Length} RGBA bytes; expected {width}x{height}.");
                }
                var errors = new List<string>();
                foreach (KeyValuePair<string, byte[]> other in resized) {
                    if (other.Key == lane.Key) continue;
                    double error = CalculateMeanAbsoluteError(other.Value, lane.Value);
                    if (error > ResizeMaximumMeanAbsoluteError) {
                        throw new InvalidOperationException(
                            $"{lane.Key} resize differs from {other.Key} by mean absolute channel error {Format(error, 3)}; limit {Format(ResizeMaximumMeanAbsoluteError, 1)}.");
                    }
                    errors.Add($"{other.Key} {Format(error, 3)}");
                }
                return $"Resize {officeImage.Width}x{officeImage.Height} to {width}x{height}; RGBA mean absolute error (limit {Format(ResizeMaximumMeanAbsoluteError, 1)}) versus " + string.Join(", ", errors);
            });
        }
    }

    /// <summary>
    /// Each encoder's output must identify as the format and source dimensions and decode exactly to the source
    /// pixels in another engine's decoder: OfficeIMO output through Magick.NET, Magick.NET output through OfficeIMO, and
    /// OneImage output (where it runs) through both.
    /// </summary>
    internal static void ValidateLosslessEncode(string operation, OfficeRasterImage source, byte[] officeEncoded, byte[] magickEncoded, OfficeImageFormat format, byte[]? oneImageEncoded = null) {
        byte[] expected = source.GetPixels();
        Record(OfficeImo, "Encode", () => {
            ImageBenchmarkCorpus.AssertIdentified(officeEncoded, format, source.Width, source.Height, operation + " OfficeIMO encode");
            AssertPixels(operation + " Magick.NET decode of OfficeIMO output", expected, ImageComparisonAdapters.DecodeMagickRgba(officeEncoded));
            return $"{operation} {source.Width}x{source.Height}, {officeEncoded.Length} bytes; Magick.NET decode equals source RGBA";
        });
        Record(Magick, "Encode", () => {
            ImageBenchmarkCorpus.AssertIdentified(magickEncoded, format, source.Width, source.Height, operation + " Magick.NET encode");
            if (!OfficeRasterImageDecoder.TryDecode(magickEncoded, out OfficeRasterImage? decoded) || decoded == null) {
                throw new InvalidOperationException(operation + " OfficeIMO could not decode Magick.NET output.");
            }
            AssertPixels(operation + " OfficeIMO decode of Magick.NET output", expected, decoded.GetPixels());
            return $"{operation} {source.Width}x{source.Height}, {magickEncoded.Length} bytes; OfficeIMO decode equals source RGBA";
        });
        if (oneImageEncoded != null) {
            Record(OneImageLibrary, "Encode", () => {
                ImageBenchmarkCorpus.AssertIdentified(oneImageEncoded, format, source.Width, source.Height, operation + " OneImage encode");
                AssertPixels(operation + " OfficeIMO decode of OneImage output", expected, ImageComparisonAdapters.DecodeOfficeImoRgba(oneImageEncoded));
                AssertPixels(operation + " Magick.NET decode of OneImage output", expected, ImageComparisonAdapters.DecodeMagickRgba(oneImageEncoded));
                return $"{operation} {source.Width}x{source.Height}, {oneImageEncoded.Length} bytes; OfficeIMO and Magick.NET decode equal source RGBA";
            });
        }
    }

    /// <summary>Both decoders must decode the externally (Magick.NET/libtiff, libwebp) written file to exactly the source pixels.</summary>
    internal static void ValidateLosslessDecode(string operation, OfficeRasterImage source, byte[] externalEncoded, OfficeImageFormat format) {
        byte[] expected = source.GetPixels();
        ImageBenchmarkCorpus.AssertIdentified(externalEncoded, format, source.Width, source.Height, operation + " external file");
        Record(OfficeImo, "DecodeExternal", () => {
            AssertPixels(operation + " OfficeIMO decode of the external file", expected, ImageComparisonAdapters.DecodeOfficeImoRgba(externalEncoded));
            return $"{operation} {source.Width}x{source.Height}; decoded RGBA equals the source the file was written from";
        });
        Record(Magick, "DecodeExternal", () => {
            AssertPixels(operation + " Magick.NET decode of the external file", expected, ImageComparisonAdapters.DecodeMagickRgba(externalEncoded));
            return $"{operation} {source.Width}x{source.Height}; decoded RGBA equals the source the file was written from";
        });
#if NET10_0_OR_GREATER
        Record(OneImageLibrary, "DecodeExternal", () => {
            AssertPixels(operation + " OneImage decode of the external file", expected, OneImageComparisonAdapters.DecodeRgba(externalEncoded));
            return $"{operation} {source.Width}x{source.Height}; decoded RGBA equals the source the file was written from";
        });
#endif
    }

    private static string Record(string library, string caseName, Func<string> check) {
        string detail;
        try {
            detail = check();
        } catch (Exception exception) {
            BenchmarkOutputValidationLog.Failed(library, exception.Message, caseName);
            throw;
        }
        BenchmarkOutputValidationLog.Passed(library, detail, caseName);
        return detail;
    }

    private static string Format(double value, int decimals) =>
        value.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

    private static OfficeRasterImage CreateOpaquePattern(int width, int height) {
        OfficeRasterImage source = ImageBenchmarkCorpus.CreatePattern(width, height);
        var opaque = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                OfficeColor pixel = source.GetPixel(x, y);
                opaque.SetPixel(x, y, OfficeColor.FromRgba(pixel.R, pixel.G, pixel.B, 255));
            }
        }
        return opaque;
    }

    internal static void AssertPixels(string engine, byte[] expected, byte[] actual) {
        if (!expected.AsSpan().SequenceEqual(actual)) {
            throw new InvalidOperationException(engine + " did not produce the expected RGBA pixels.");
        }
    }

    private static double CalculateMeanAbsoluteError(byte[] expected, byte[] actual) {
        if (expected.Length != actual.Length) {
            throw new InvalidOperationException("The image engine did not produce the expected pixel count.");
        }
        long absoluteError = 0L;
        for (int index = 0; index < expected.Length; index++) {
            absoluteError += Math.Abs(expected[index] - actual[index]);
        }
        return absoluteError / (double)expected.Length;
    }
}
