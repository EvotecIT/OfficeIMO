using OfficeIMO.Drawing;
using System;
using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private const long MaximumIccPngPixels = 4_000_000L;

    private static bool TryNormalizeIccPng(
        byte[] source,
        OfficeImageInfo sourceInfo,
        byte[] profile,
        CancellationToken cancellationToken,
        out byte[] normalized,
        out string? reason) {
        normalized = Array.Empty<byte>();
        reason = null;
        if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                sourceInfo, MaximumIccPngPixels, out reason)) return false;
        long pixels = (long)sourceInfo.Width * sourceInfo.Height;
        // Decoder RGBA, packed RGB, converted RGBA, and the PNG encoder coexist briefly.
        // Leave room for the ICC parser and encoder's transient buffers before decoding.
        long estimatedPeak = source.LongLength + (long)profile.Length * 33L + 4096L + pixels * 32L;
        if (estimatedPeak > OfficeRasterGuards.MaximumDecodedBytes) {
            reason = "The ICC-tagged PNG exceeds the bounded color-normalization memory limit.";
            return false;
        }
        // PDF embeds the static IDAT image of an APNG, which can differ from its first fdAT frame.
        if (!OfficePngReader.TryDecode(source, cancellationToken, out OfficeRasterImage? decoded) || decoded == null ||
            decoded.Width != sourceInfo.Width || decoded.Height != sourceInfo.Height) {
            reason = "The ICC-tagged PNG pixels could not be decoded safely.";
            return false;
        }

        if (!OfficeIccColorProfile.TryCreate(profile, out OfficeIccColorProfile? colorProfile) ||
            colorProfile == null) {
            reason = "The embedded PNG ICC profile has no supported color transform.";
            return false;
        }
        int pngColorType = source[25]; // The validated PNG's IHDR color-type byte.
        int channels = pngColorType == 0 || pngColorType == 4 ? 1 : 3;
        if (colorProfile.ComponentCount != channels) {
            reason = "The embedded PNG ICC profile does not match its color type.";
            return false;
        }
        byte[] rgba = decoded.PixelBuffer;
        var samples = new byte[checked((int)(pixels * channels))];
        for (int pixel = 0; pixel < pixels; pixel++) {
            if ((pixel & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int sourceOffset = pixel * 4;
            int sampleOffset = pixel * channels;
            samples[sampleOffset] = rgba[sourceOffset];
            if (channels == 3) {
                samples[sampleOffset + 1] = rgba[sourceOffset + 1];
                samples[sampleOffset + 2] = rgba[sourceOffset + 2];
            }
        }

        OfficeIccRasterConversionStatus status = OfficeIccRasterConverter.TryConvertToSrgb(
            samples, decoded.Width, decoded.Height, profile,
            new OfficeIccRasterConversionOptions {
                MaximumPixels = MaximumIccPngPixels,
                CancellationToken = cancellationToken
            }, out OfficeRasterImage? converted);
        if (status != OfficeIccRasterConversionStatus.Converted || converted == null) {
            reason = "The embedded PNG ICC profile could not normalize its color samples to sRGB (" + status + ").";
            return false;
        }

        byte[] convertedPixels = converted.PixelBuffer;
        for (int pixel = 0; pixel < pixels; pixel++) {
            if ((pixel & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            convertedPixels[pixel * 4 + 3] = rgba[pixel * 4 + 3];
        }
        normalized = OfficePngWriter.Encode(converted, cancellationToken);
        return true;
    }
}
