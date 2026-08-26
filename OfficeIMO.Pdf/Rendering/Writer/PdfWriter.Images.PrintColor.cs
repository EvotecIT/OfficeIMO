using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryConvertImageStreamToCmyk(
        byte[] sourceBytes,
        PdfImageStream image,
        PdfPrintColorTransform transform,
        bool flattenTransparency,
        PdfColor transparencyBackground,
        System.Threading.CancellationToken cancellationToken,
        out string? unsupportedReason) {
        cancellationToken.ThrowIfCancellationRequested();
        unsupportedReason = null;
        long declaredPixelCount = (long)image.PixelWidth * image.PixelHeight;
        if (declaredPixelCount <= 0 || declaredPixelCount > MaxPngPixelCount) {
            unsupportedReason = "Raster image dimensions exceed the bounded PDF/X CMYK conversion limit.";
            return false;
        }

        if (!OfficeRasterImageDecoder.TryDecode(
                sourceBytes,
                options: null,
                out OfficeRasterImage? raster,
                out OfficeRasterDecodeInfo decodeInfo) || raster == null) {
            unsupportedReason = "The raster image decoder could not produce RGBA pixels for CMYK conversion.";
            return false;
        }

        long pixelCount = (long)raster.Width * raster.Height;
        if (pixelCount <= 0 || pixelCount > MaxPngPixelCount) {
            unsupportedReason = "Raster image dimensions exceed the bounded PDF/X CMYK conversion limit.";
            return false;
        }

        // The metadata inspector accounts for the encoded source separately. Report only
        // the decoded raster retained beside it so the shared aggregate budget is additive.
        long retainedBeforeMetadata = checked(pixelCount * 4L + 24L);
        OfficeImageMetadataSnapshot metadata = OfficeImageMetadataInspector.Inspect(
            sourceBytes,
            decodeInfo.Format,
            retainedBeforeMetadata);
        long baseRowLengthValue = checked(1L + raster.Width * 4L);
        long cmykLengthValue = checked(baseRowLengthValue * raster.Height);
        long alphaLengthValue = flattenTransparency ? 0L : checked((1L + raster.Width) * raster.Height);
        if (baseRowLengthValue > int.MaxValue ||
            cmykLengthValue > int.MaxValue ||
            alphaLengthValue > int.MaxValue) {
            unsupportedReason = "Raster image conversion exceeds the bounded PDF/X managed-memory limit.";
            return false;
        }
        OfficeIccColorProfile? sourceProfile = null;
        if ((metadata.Kinds & OfficeImageMetadataKinds.Icc) != 0) {
            if (metadata.Icc == null ||
                !TryGetIccParseAllocationUpperBound(metadata.Icc.LongLength, out long profileParseUpperBound) ||
                !IsPdfXImageWorkingSetWithinLimit(
                    sourceBytes.LongLength,
                    raster.PixelBuffer.LongLength,
                    metadata.Icc.LongLength,
                    profileParseUpperBound,
                    cmykLengthValue,
                    alphaLengthValue)) {
                unsupportedReason = "The raster image ICC profile exceeds the bounded PDF/X managed-memory limit before parsing.";
                return false;
            }
            if (!OfficeIccColorProfile.TryCreate(metadata.Icc, out sourceProfile) ||
                sourceProfile == null ||
                sourceProfile.ComponentCount != 3) {
                unsupportedReason = "The raster image carries an embedded ICC profile that cannot be normalized to sRGB before PDF/X CMYK conversion.";
                return false;
            }
        }

        if (!IsPdfXImageWorkingSetWithinLimit(
                sourceBytes.LongLength,
                raster.PixelBuffer.LongLength,
                metadata.Icc?.LongLength ?? 0L,
                sourceProfile?.RetainedByteCount ?? 0L,
                cmykLengthValue,
                alphaLengthValue)) {
            unsupportedReason = "Raster image conversion exceeds the bounded PDF/X managed-memory limit.";
            return false;
        }

        byte[] pixels = raster.PixelBuffer;
        int baseRowLength = checked((int)baseRowLengthValue);
        byte[] cmykRows = new byte[checked((int)cmykLengthValue)];
        byte[]? alphaRows = flattenTransparency ? null : new byte[checked((int)alphaLengthValue)];
        PdfColor background = transparencyBackground;
        var components = new double[4];
        var sourceComponents = sourceProfile == null ? null : new double[3];
        bool hasTransparency = false;

        for (int row = 0; row < raster.Height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceRow = row * raster.Width * 4;
            int targetRow = row * baseRowLength;
            int alphaRow = row * (1 + raster.Width);
            cmykRows[targetRow] = 0;
            if (alphaRows != null) alphaRows[alphaRow] = 0;
            for (int column = 0; column < raster.Width; column++) {
                if ((column & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                int source = sourceRow + column * 4;
                byte alpha = pixels[source + 3];
                hasTransparency |= alpha != 255;
                OfficeColor color;
                if (sourceProfile != null) {
                    sourceComponents![0] = pixels[source] / 255D;
                    sourceComponents[1] = pixels[source + 1] / 255D;
                    sourceComponents[2] = pixels[source + 2] / 255D;
                    if (!sourceProfile.TryConvert(sourceComponents, transform.RenderingIntent, out color)) {
                        unsupportedReason = "The raster image ICC profile could not normalize a source pixel to sRGB before PDF/X CMYK conversion.";
                        return false;
                    }
                } else {
                    color = OfficeColor.FromRgb(pixels[source], pixels[source + 1], pixels[source + 2]);
                }
                if (flattenTransparency && alpha != 255) {
                    double opacity = alpha / 255D;
                    color = OfficeColor.FromRgb(
                        Composite(color.R, background.R, opacity),
                        Composite(color.G, background.G, opacity),
                        Composite(color.B, background.B, opacity));
                }

                transform.Convert(color, components);
                int target = targetRow + 1 + column * 4;
                cmykRows[target] = ToComponentByte(components[0]);
                cmykRows[target + 1] = ToComponentByte(components[1]);
                cmykRows[target + 2] = ToComponentByte(components[2]);
                cmykRows[target + 3] = ToComponentByte(components[3]);
                if (alphaRows != null) alphaRows[alphaRow + 1 + column] = alpha;
            }
        }

        image.Data = DeflateZlib(cmykRows);
        image.PixelWidth = raster.Width;
        image.PixelHeight = raster.Height;
        image.DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceCMYK", 4, raster.Width);
        image.SoftMask = hasTransparency && alphaRows != null
            ? new PdfImageStream {
                Data = DeflateZlib(alphaRows),
                PixelWidth = raster.Width,
                PixelHeight = raster.Height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, raster.Width)
            }
            : null;
        return true;
    }

    private static byte Composite(byte foreground, double background, double opacity) =>
        (byte)Math.Round((foreground * opacity) + (background * 255D * (1D - opacity)));

    internal static bool IsPdfXImageWorkingSetWithinLimit(
        long sourceBytes,
        long rasterBytes,
        long profileBytes,
        long profileTransformBytes,
        long cmykBytes,
        long alphaBytes) {
        try {
            long baseline = checked(
                sourceBytes + 24L +
                rasterBytes + 24L +
                (profileBytes == 0L ? 0L : profileBytes + 24L) +
                profileTransformBytes +
                cmykBytes + 24L +
                (alphaBytes == 0L ? 0L : alphaBytes + 24L));
            long cmykCompressedBound = GetDeflateZlibMaximumLength(cmykBytes);
            long cmykCompressionPeak = checked(baseline + cmykCompressedBound * 3L + 48L);
            long alphaCompressionPeak = alphaBytes == 0L
                ? 0L
                : checked(
                    baseline + cmykCompressedBound + 24L +
                    GetDeflateZlibMaximumLength(alphaBytes) * 3L + 48L);
            return Math.Max(cmykCompressionPeak, alphaCompressionPeak) <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool TryGetIccParseAllocationUpperBound(long profileBytes, out long upperBound) {
        upperBound = 0L;
        if (profileBytes <= 0L) return false;
        try {
            // Supported ICC LUT and curve payloads can expand from compact byte/ushort samples
            // into doubles and multiple intent transforms. Bound the parser before it materializes
            // those structures; the post-parse check below uses the exact retained count.
            upperBound = checked(1_048_576L + profileBytes * 64L);
            return upperBound <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static long GetDeflateZlibMaximumLength(long length) => checked(
        length + (length >> 12) + (length >> 14) + (length >> 25) + 13L);

    private static byte ToComponentByte(double value) =>
        (byte)Math.Round(Math.Max(0D, Math.Min(1D, value)) * 255D);
}
