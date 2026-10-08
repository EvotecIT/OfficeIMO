using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Encodes multi-resolution Windows icons using independently compressed PNG entries.</summary>
public static class OfficeIconEncoder {
    /// <summary>Creates an ICO container from one to 256 images, each between one and 256 pixels on each axis.</summary>
    public static byte[] Encode(IReadOnlyList<OfficeRasterImage> images, OfficeRasterEncodingOptions? options = null, CancellationToken cancellationToken = default) {
        if (images == null) throw new ArgumentNullException(nameof(images));
        if (images.Count < 1 || images.Count > 256) throw new ArgumentOutOfRangeException(nameof(images), "An icon must contain between one and 256 images.");
        var payloads = new byte[images.Count][]; long total = 6L + images.Count * 16L; long rasterBytes = 0;
        for (int i = 0; i < images.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested(); OfficeRasterImage image = images[i] ?? throw new ArgumentException("Icon images cannot be null.", nameof(images));
            if (image.Width > 256 || image.Height > 256) throw new ArgumentOutOfRangeException(nameof(images), "Icon dimensions must not exceed 256 pixels.");
            rasterBytes = checked(rasterBytes + image.Width * (long)image.Height * 4);
            payloads[i] = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options, OfficeRasterGuards.MaximumEncodedBytes - total, cancellationToken); total = checked(total + payloads[i].Length);
        }
        int size = OfficeRasterGuards.EnsureOutputBytes(total, "Icon output exceeds the encoded-size limit.");
        if (total * 2 + rasterBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Icon encoding exceeds the managed working-set limit.", nameof(images));
        var result = new byte[size]; OfficeExifProfileCodec.Write(result, 2, 1, 2, true); OfficeExifProfileCodec.Write(result, 4, (uint)images.Count, 2, true); int cursor = 6 + images.Count * 16;
        for (int i = 0; i < images.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested(); int entry = 6 + i * 16; result[entry] = (byte)images[i].Width; result[entry + 1] = (byte)images[i].Height;
            OfficeExifProfileCodec.Write(result, entry + 4, 1, 2, true); OfficeExifProfileCodec.Write(result, entry + 6, 32, 2, true); OfficeExifProfileCodec.Write(result, entry + 8, (uint)payloads[i].Length, 4, true); OfficeExifProfileCodec.Write(result, entry + 12, (uint)cursor, 4, true);
            Buffer.BlockCopy(payloads[i], 0, result, cursor, payloads[i].Length); cursor += payloads[i].Length;
        }
        return result;
    }
}

internal static class OfficeIconDecoder {
    internal static bool TryInspect(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterContainerInfo? container) {
        container = null;
        if (!OfficeImageReader.TryIdentifyByContent(bytes, null, options.CancellationToken, out OfficeImageInfo info) || info.Format != OfficeImageFormat.Icon) return false;
        int count = U16(bytes, 4); if (count > 1024) return false;
        var frames = new OfficeRasterFrameInfo[count]; long totalPixels = 0;
        for (int i = 0; i < count; i++) {
            int entry = 6 + i * 16; int width = bytes[entry] == 0 ? 256 : bytes[entry]; int height = bytes[entry + 1] == 0 ? 256 : bytes[entry + 1]; totalPixels += (long)width * height;
            if (totalPixels > options.MaximumInspectionWorkPixels || !OfficeRasterImageDecoder.IsWithinPixelLimit(width, height, options.MaximumDecodedPixels)) return false;
            frames[i] = new OfficeRasterFrameInfo(i, OfficeRasterFrameKind.Image, width, height, 0, 0, TimeSpan.Zero, OfficeRasterFrameDisposal.None, OfficeRasterFrameBlend.Source, i == 0);
        }
        if (!OfficeImageReader.TryValidateContent(bytes, null, options.CancellationToken, out _)) return false;
        container = new OfficeRasterContainerInfo(OfficeImageFormat.Icon, info.Width, info.Height, frames, 1, OfficeColor.Transparent); return true;
    }
    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image) {
        image = null;
        try {
            int entry = checked(6 + options.FrameIndex * 16); int length = checked((int)U32(bytes, entry + 8)); int offset = checked((int)U32(bytes, entry + 12));
            if (length > options.MaximumEncodedBytes || (long)length + bytes.Length + options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            var payload = new byte[length]; Buffer.BlockCopy(bytes, offset, payload, 0, length);
            if (length >= 8 && payload[0] == 137 && payload[1] == 80 && payload[2] == 78 && payload[3] == 71) {
                var nested = options.WithAdditionalRetainedManagedBytes(bytes.LongLength); nested.FrameIndex = 0;
                return OfficeRasterImageDecoder.TryDecode(payload, nested, out image, out _);
            }
            int header = checked((int)U32(payload, 0)); int width = header == 12 ? U16(payload, 4) : checked((int)U32(payload, 4)); int height = (header == 12 ? U16(payload, 6) : checked((int)U32(payload, 8))) / 2;
            int depth = U16(payload, header == 12 ? 10 : 14); int compression = header == 12 ? 0 : checked((int)U32(payload, 16));
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, options.MaximumDecodedPixels, out int pixels) || pixels * 4L + payload.Length + bytes.Length + options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int paletteCount = depth <= 8 ? 1 << depth : 0; if (header >= 40 && U32(payload, 32) != 0) paletteCount = checked((int)U32(payload, 32));
            int externalMasks = header == 40 && compression == 3 ? 12 : header == 40 && compression == 6 ? 16 : 0;
            int paletteOffset = header + externalMasks; int pixelOffset = paletteOffset + paletteCount * (header == 12 ? 3 : 4); int stride = ((width * depth + 31) / 32) * 4; int maskStride = ((width + 31) / 32) * 4; int maskOffset = pixelOffset + stride * height; bool hasMask = payload.Length >= maskOffset + maskStride * height;
            uint redMask = depth == 16 ? 0x7C00U : 0x00FF0000U, greenMask = depth == 16 ? 0x03E0U : 0x0000FF00U, blueMask = depth == 16 ? 0x001FU : 0x000000FFU, alphaMask = 0;
            if (compression == 3 || compression == 6) { redMask = U32(payload, 40); greenMask = U32(payload, 44); blueMask = U32(payload, 48); if (compression == 6 || header >= 56) alphaMask = U32(payload, 52); }
            bool useStoredAlpha = alphaMask != 0;
            if (depth == 32 && compression == 0) { for (int y = 0; y < height && !useStoredAlpha; y++) for (int x = 0; x < width; x++) if (payload[pixelOffset + y * stride + x * 4 + 3] != 0) { useStoredAlpha = true; break; } alphaMask = 0xFF000000U; }
            var result = new OfficeRasterImage(width, height);
            for (int y = 0; y < height; y++) {
                options.CancellationToken.ThrowIfCancellationRequested(); int srcY = height - 1 - y;
                for (int x = 0; x < width; x++) {
                    int at = pixelOffset + srcY * stride; byte red, green, blue, alpha = 255;
                    if (depth <= 8) { int index = depth == 8 ? payload[at + x] : depth == 4 ? (payload[at + x / 2] >> (x % 2 == 0 ? 4 : 0)) & 15 : (payload[at + x / 8] >> (7 - x % 8)) & 1; int color = paletteOffset + index * (header == 12 ? 3 : 4); blue = payload[color]; green = payload[color + 1]; red = payload[color + 2]; }
                    else if (depth == 24) { at += x * 3; blue = payload[at]; green = payload[at + 1]; red = payload[at + 2]; }
                    else { at += x * (depth / 8); uint packed = depth == 16 ? (uint)U16(payload, at) : U32(payload, at); red = Channel(packed, redMask); green = Channel(packed, greenMask); blue = Channel(packed, blueMask); if (useStoredAlpha) alpha = Channel(packed, alphaMask); }
                    if (hasMask && (payload[maskOffset + srcY * maskStride + x / 8] & (1 << (7 - x % 8))) != 0) alpha = 0;
                    result.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, alpha));
                }
            }
            image = result; return true;
        } catch (OperationCanceledException) { throw; } catch (Exception ex) when (ex is FormatException || ex is OverflowException || ex is ArgumentException || ex is IndexOutOfRangeException) { return false; }
    }
    private static byte Channel(uint packed, uint mask) { if (mask == 0) return 255; int shift = 0; while ((mask & 1) == 0) { mask >>= 1; shift++; } return (byte)(((ulong)((packed >> shift) & mask) * 255 + mask / 2) / mask); }
    private static int U16(byte[] bytes, int offset) => bytes[offset] | bytes[offset + 1] << 8;
    private static uint U32(byte[] bytes, int offset) => (uint)OfficeExifProfileCodec.Read(bytes, offset, 4, true);
}
