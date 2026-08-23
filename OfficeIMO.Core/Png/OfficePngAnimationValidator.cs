using System;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Validates APNG sequencing, frame geometry, and secondary-frame scanline payloads.</summary>
internal static class OfficePngAnimationValidator {
    private const int SignatureLength = 8;

    internal static bool TryValidateAdditionalFrames(
        byte[] bytes,
        CancellationToken cancellationToken = default) {
        try {
            int canvasWidth = 0;
            int canvasHeight = 0;
            int bitDepth = 0;
            int colorType = 0;
            int interlaceMethod = 0;
            byte[]? palette = null;
            bool seenAnimationControl = false;
            bool seenImageData = false;
            int declaredFrameCount = 0;
            int frameControlCount = 0;
            long decodedFramePixels = 0;
            uint expectedSequence = 0;
            FramePayload? currentFrame = null;

            int offset = SignatureLength;
            while (offset + 12 <= bytes.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int length = ReadBigEndianInt32(bytes, offset);
                int dataOffset = offset + 8;
                string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
                switch (type) {
                    case "IHDR":
                        canvasWidth = ReadBigEndianInt32(bytes, dataOffset);
                        canvasHeight = ReadBigEndianInt32(bytes, dataOffset + 4);
                        bitDepth = bytes[dataOffset + 8];
                        colorType = bytes[dataOffset + 9];
                        interlaceMethod = bytes[dataOffset + 12];
                        break;
                    case "PLTE":
                        palette = new byte[length];
                        Buffer.BlockCopy(bytes, dataOffset, palette, 0, length);
                        break;
                    case "acTL":
                        seenAnimationControl = true;
                        declaredFrameCount = ReadBigEndianInt32(bytes, dataOffset);
                        break;
                    case "fcTL":
                        if (!seenAnimationControl || length != 26 ||
                            (currentFrame != null && currentFrame.UsesDefaultImageData && !seenImageData) ||
                            !TryFinishFrame(
                                currentFrame, bitDepth, colorType, interlaceMethod, palette, cancellationToken)) {
                            return false;
                        }
                        uint frameSequence = ReadBigEndianUInt32(bytes, dataOffset);
                        if (frameSequence != expectedSequence++) return false;
                        int width = ReadBigEndianInt32(bytes, dataOffset + 4);
                        int height = ReadBigEndianInt32(bytes, dataOffset + 8);
                        int x = ReadBigEndianInt32(bytes, dataOffset + 12);
                        int y = ReadBigEndianInt32(bytes, dataOffset + 16);
                        bool isFirstAnimationFrame = frameControlCount == 0;
                        if (!HasValidFrameBounds(width, height, x, y, canvasWidth, canvasHeight, out int framePixels) ||
                            decodedFramePixels > OfficeRasterGuards.MaximumPixels - framePixels ||
                            bytes[dataOffset + 24] > 2 || bytes[dataOffset + 25] > 1 ||
                            isFirstAnimationFrame && bytes[dataOffset + 24] == 2) {
                            return false;
                        }
                        decodedFramePixels += framePixels;

                        bool usesDefaultImageData = frameControlCount == 0 && !seenImageData;
                        if (usesDefaultImageData &&
                            (x != 0 || y != 0 || width != canvasWidth || height != canvasHeight)) {
                            return false;
                        }
                        currentFrame = new FramePayload(width, height, usesDefaultImageData);
                        frameControlCount++;
                        break;
                    case "IDAT":
                        seenImageData = true;
                        break;
                    case "fdAT":
                        if (!seenAnimationControl || !seenImageData || length < 4 ||
                            currentFrame == null || currentFrame.UsesDefaultImageData) {
                            return false;
                        }
                        uint dataSequence = ReadBigEndianUInt32(bytes, dataOffset);
                        if (dataSequence != expectedSequence++) return false;
                        WritePayload(
                            currentFrame.Compressed, bytes, dataOffset + 4, length - 4, cancellationToken);
                        break;
                    case "IEND":
                        if (!seenAnimationControl) return true;
                        return declaredFrameCount > 0 &&
                               frameControlCount == declaredFrameCount &&
                               TryFinishFrame(
                                   currentFrame, bitDepth, colorType, interlaceMethod, palette, cancellationToken);
                }

                offset += 12 + length;
            }
            return false;
        } catch (Exception exception) when (
            exception is ArgumentException ||
            exception is FormatException ||
            exception is IOException ||
            exception is OverflowException) {
            return false;
        }
    }

    private static bool TryFinishFrame(
        FramePayload? frame,
        int bitDepth,
        int colorType,
        int interlaceMethod,
        byte[]? palette,
        CancellationToken cancellationToken) {
        if (frame == null || frame.UsesDefaultImageData) return true;
        return OfficePngReader.TryValidateCompressedPayload(
            frame.Compressed.ToArray(),
            frame.Width,
            frame.Height,
            bitDepth,
            colorType,
            interlaceMethod,
            palette,
            cancellationToken);
    }

    private static void WritePayload(
        Stream destination,
        byte[] bytes,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        const int copyChunkBytes = 64 * 1024;
        int remaining = count;
        while (remaining > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int copy = Math.Min(remaining, copyChunkBytes);
            destination.Write(bytes, offset, copy);
            offset += copy;
            remaining -= copy;
        }
    }

    private static bool HasValidFrameBounds(
        int width,
        int height,
        int x,
        int y,
        int canvasWidth,
        int canvasHeight,
        out int framePixels) {
        framePixels = 0;
        return width > 0 && height > 0 && x >= 0 && y >= 0 &&
               (long)x + width <= canvasWidth &&
               (long)y + height <= canvasHeight &&
               OfficeRasterGuards.TryEnsurePixelCount(width, height, out framePixels);
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];

    private sealed class FramePayload {
        internal FramePayload(int width, int height, bool usesDefaultImageData) {
            Width = width;
            Height = height;
            UsesDefaultImageData = usesDefaultImageData;
        }

        internal int Width { get; }

        internal int Height { get; }

        internal bool UsesDefaultImageData { get; }

        internal MemoryStream Compressed { get; } = new();
    }
}
