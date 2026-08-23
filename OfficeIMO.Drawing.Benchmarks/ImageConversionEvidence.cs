namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Validates the supported static raster conversion matrix and animation-loss boundary.</summary>
internal static class ImageConversionEvidence {
    private sealed record ConversionSource(string Name, byte[] Bytes);

    private static readonly OfficeImageFormat[] OutputFormats = {
        OfficeImageFormat.Png,
        OfficeImageFormat.Jpeg,
        OfficeImageFormat.Tiff,
        OfficeImageFormat.Webp
    };

    internal static void Validate(TextWriter writer) {
        OfficeRasterImage pattern = ImageBenchmarkCorpus.CreatePattern(64, 48);
        var sourceOptions = new OfficeRasterEncodingOptions {
            DpiX = 144D,
            DpiY = 120D
        };
        var sources = new[] {
            new ConversionSource("PNG", OfficeRasterImageEncoder.Encode(pattern, OfficeImageExportFormat.Png, sourceOptions)),
            new ConversionSource("JPEG", OfficeRasterImageEncoder.Encode(pattern, OfficeImageExportFormat.Jpeg, sourceOptions)),
            new ConversionSource("TIFF", OfficeRasterImageEncoder.Encode(pattern, OfficeImageExportFormat.Tiff, sourceOptions)),
            new ConversionSource("WebP", OfficeRasterImageEncoder.Encode(pattern, OfficeImageExportFormat.Webp, sourceOptions)),
            new ConversionSource("BMP", ImageBenchmarkCorpus.CreateBmp24(64, 48)),
            new ConversionSource("GIF", Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="))
        };

        writer.WriteLine();
        writer.WriteLine("Static conversion matrix (source -> output, validated dimensions/DPI/pixels):");
        writer.WriteLine("Source     Output       Dimensions      Bytes       MAE");
        foreach (ConversionSource source in sources) {
            OfficeImageInfo sourceInfo = OfficeImageReader.Identify(source.Bytes);
            OfficeRasterImage decodedSource = ImageBenchmarkCorpus.Decode(source.Bytes, source.Name + " conversion source");
            int targetWidth = Math.Max(1, decodedSource.Width / 2);
            int targetHeight = Math.Max(1, decodedSource.Height / 2);
            OfficeRasterImage expected = targetWidth == decodedSource.Width && targetHeight == decodedSource.Height
                ? decodedSource
                : OfficeRasterResampler.Resize(decodedSource, targetWidth, targetHeight);

            foreach (OfficeImageFormat outputFormat in OutputFormats) {
                OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                    source.Bytes,
                    new OfficeImageOptimizationRequest(targetWidth, targetHeight) {
                        PreserveAspectRatio = false,
                        OutputFormat = outputFormat,
                        KeepOriginalWhenNotSmaller = false,
                        JpegSubsampling = OfficeJpegSubsampling.Y420,
                        JpegOptimizeHuffman = true,
                        TiffCompression = OfficeTiffCompression.Deflate
                    });
                if (result.Status != OfficeImageOptimizationStatus.Optimized) {
                    throw new InvalidOperationException(
                        $"{source.Name} to {outputFormat} returned {result.Status} instead of Optimized.");
                }

                byte[] resultBytes = result.Bytes;
                OfficeImageInfo actualInfo = OfficeImageReader.Identify(resultBytes);
                if (actualInfo.Format != outputFormat ||
                    actualInfo.Width != targetWidth ||
                    actualInfo.Height != targetHeight ||
                    Math.Abs(actualInfo.DpiX - sourceInfo.DpiX) > 0.05D ||
                    Math.Abs(actualInfo.DpiY - sourceInfo.DpiY) > 0.05D ||
                    actualInfo.DpiX != result.Final.DpiX ||
                    actualInfo.DpiY != result.Final.DpiY) {
                    throw new InvalidOperationException(
                        $"{source.Name} to {outputFormat} did not preserve its encoded format, dimensions, or physical resolution contract.");
                }

                OfficeRasterImage actual = ImageBenchmarkCorpus.Decode(resultBytes, source.Name + " to " + outputFormat);
                double meanAbsoluteError;
                if (outputFormat == OfficeImageFormat.Jpeg) {
                    byte[] flattened = ImageEncodingEvidence.FlattenAgainstWhite(expected.GetPixels());
                    (meanAbsoluteError, double psnr) = ImageEncodingEvidence.MeasureRgbFidelity(flattened, actual.GetPixels());
                    if (meanAbsoluteError > 40D || psnr < 15D) {
                        throw new InvalidOperationException(
                            $"{source.Name} to JPEG fidelity was outside the validation envelope: MAE {meanAbsoluteError:F3}, PSNR {psnr:F2} dB.");
                    }
                } else {
                    byte[] expectedPixels = expected.GetPixels();
                    byte[] actualPixels = actual.GetPixels();
                    if (!expectedPixels.AsSpan().SequenceEqual(actualPixels)) {
                        throw new InvalidOperationException(source.Name + " to " + outputFormat + " was not lossless.");
                    }
                    meanAbsoluteError = 0D;
                }

                writer.WriteLine(
                    $"{source.Name,-10} {outputFormat,-10} {targetWidth,4}x{targetHeight,-4} {resultBytes.Length,10:N0} {meanAbsoluteError,9:F3}");
            }
        }

        byte[] animation = ImageBenchmarkCorpus.Animation.ReadBytes();
        OfficeImageOptimizationResult animated = OfficeImageOptimizer.Optimize(
            animation,
            new OfficeImageOptimizationRequest(110, 69) {
                KeepOriginalWhenNotSmaller = false
            });
        if (animated.Status != OfficeImageOptimizationStatus.DecodeFailed || animated.Changed) {
            throw new InvalidOperationException("Animated GIF optimization did not preserve the no-implicit-frame-loss boundary.");
        }
        writer.WriteLine("Animated GIF input was rejected without silently selecting one frame.");
    }
}
