using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Validates APNG sequencing, frame geometry, and secondary-frame scanline payloads.</summary>
internal static class OfficePngAnimationValidator {
    private const int SignatureLength = 8;

    internal static bool TryValidateStructure(
        byte[] bytes,
        CancellationToken cancellationToken = default) =>
        TryValidate(bytes, validateCompressedPayloads: false, cancellationToken);

    internal static bool TryValidateAdditionalFrames(
        byte[] bytes,
        CancellationToken cancellationToken = default) =>
        TryValidate(bytes, validateCompressedPayloads: true, cancellationToken);

    private static bool TryValidate(
        byte[] bytes,
        bool validateCompressedPayloads,
        CancellationToken cancellationToken) {
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
                        if (seenAnimationControl || seenImageData || length != 8) return false;
                        uint encodedFrameCount = ReadBigEndianUInt32(bytes, dataOffset);
                        if (encodedFrameCount == 0 || encodedFrameCount > int.MaxValue) return false;
                        seenAnimationControl = true;
                        declaredFrameCount = (int)encodedFrameCount;
                        break;
                    case "fcTL":
                        if (length != 26 ||
                            (currentFrame != null && currentFrame.UsesDefaultImageData && !seenImageData) ||
                            !TryFinishFrame(
                                bytes, currentFrame, validateCompressedPayloads, bitDepth, colorType, interlaceMethod,
                                palette, cancellationToken)) {
                            return false;
                        }
                        uint frameSequence = ReadBigEndianUInt32(bytes, dataOffset);
                        if (frameSequence != expectedSequence++) return false;
                        int width = ReadBigEndianInt32(bytes, dataOffset + 4);
                        int height = ReadBigEndianInt32(bytes, dataOffset + 8);
                        int x = ReadBigEndianInt32(bytes, dataOffset + 12);
                        int y = ReadBigEndianInt32(bytes, dataOffset + 16);
                        if (!HasValidFrameBounds(width, height, x, y, canvasWidth, canvasHeight, out int framePixels) ||
                            (validateCompressedPayloads &&
                             decodedFramePixels > OfficeRasterGuards.MaximumPixels - framePixels) ||
                            bytes[dataOffset + 24] > 2 || bytes[dataOffset + 25] > 1) {
                            return false;
                        }
                        if (validateCompressedPayloads) decodedFramePixels += framePixels;

                        bool usesDefaultImageData = frameControlCount == 0 && !seenImageData;
                        if (usesDefaultImageData &&
                            (x != 0 || y != 0 || width != canvasWidth || height != canvasHeight)) {
                            return false;
                        }
                        currentFrame = new FramePayload(width, height, usesDefaultImageData, validateCompressedPayloads);
                        frameControlCount++;
                        break;
                    case "IDAT":
                        if (frameControlCount > 0 && !seenAnimationControl) return false;
                        seenImageData = true;
                        break;
                    case "fdAT":
                        if (!seenAnimationControl || !seenImageData || length < 4 ||
                            currentFrame == null || currentFrame.UsesDefaultImageData) {
                            return false;
                        }
                        uint dataSequence = ReadBigEndianUInt32(bytes, dataOffset);
                        if (dataSequence != expectedSequence++) return false;
                        currentFrame.SawDataChunk = true;
                        if (validateCompressedPayloads) {
                            currentFrame.AddSegment(dataOffset + 4, length - 4, bytes.Length);
                        }
                        break;
                    case "IEND":
                        if (!seenAnimationControl) return frameControlCount == 0;
                        return declaredFrameCount > 0 &&
                               frameControlCount == declaredFrameCount &&
                               TryFinishFrame(
                                   bytes, currentFrame, validateCompressedPayloads, bitDepth, colorType, interlaceMethod,
                                   palette, cancellationToken);
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
        byte[] source,
        FramePayload? frame,
        bool validateCompressedPayloads,
        int bitDepth,
        int colorType,
        int interlaceMethod,
        byte[]? palette,
        CancellationToken cancellationToken) {
        if (frame == null || frame.UsesDefaultImageData) return true;
        if (!frame.SawDataChunk) return false;
        if (!validateCompressedPayloads) return true;
        if (!OfficePngReader.TryGetValidationWorkingSetBytes(
                frame.Width, frame.Height, bitDepth, colorType, interlaceMethod, palette,
                out long validationWorkingSetBytes)) return false;
        long peakBytes = checked(
            source.Length + frame.CompressedLength + validationWorkingSetBytes +
            frame.Segments!.Count * 16L + (palette?.Length ?? 0) + 64L * 1024L);
        if (peakBytes > OfficeRasterGuards.MaximumDecodedBytes || frame.CompressedLength > int.MaxValue) return false;
        byte[] compressed = new byte[(int)frame.CompressedLength];
        int destinationOffset = 0;
        foreach (FrameSegment segment in frame.Segments) {
            CopyPayload(source, segment.Offset, compressed, destinationOffset, segment.Length, cancellationToken);
            destinationOffset += segment.Length;
        }
        return OfficePngReader.TryValidateCompressedPayload(
            compressed,
            frame.Width,
            frame.Height,
            bitDepth,
            colorType,
            interlaceMethod,
            palette,
            cancellationToken);
    }

    private static void CopyPayload(
        byte[] source,
        int sourceOffset,
        byte[] destination,
        int destinationOffset,
        int count,
        CancellationToken cancellationToken) {
        const int copyChunkBytes = 64 * 1024;
        int remaining = count;
        while (remaining > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int copy = Math.Min(remaining, copyChunkBytes);
            Buffer.BlockCopy(source, sourceOffset, destination, destinationOffset, copy);
            sourceOffset += copy;
            destinationOffset += copy;
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
        internal FramePayload(int width, int height, bool usesDefaultImageData, bool captureCompressedPayload) {
            Width = width;
            Height = height;
            UsesDefaultImageData = usesDefaultImageData;
            Segments = captureCompressedPayload ? new List<FrameSegment>() : null;
        }

        internal int Width { get; }

        internal int Height { get; }

        internal bool UsesDefaultImageData { get; }

        internal bool SawDataChunk { get; set; }

        internal long CompressedLength { get; private set; }

        internal List<FrameSegment>? Segments { get; }

        internal void AddSegment(int offset, int length, int encodedLength) {
            if (length < 0 || CompressedLength > OfficeRasterGuards.MaximumEncodedBytes - length) {
                throw new FormatException("APNG compressed frame data exceeds size limits.");
            }
            long segmentBytes = checked(((long)Segments!.Count + 1L) * 16L);
            if (encodedLength + segmentBytes > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new FormatException("APNG frame segment metadata exceeds memory limits.");
            }
            Segments.Add(new FrameSegment(offset, length));
            CompressedLength += length;
        }
    }

    private readonly struct FrameSegment {
        internal FrameSegment(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }

        internal int Length { get; }
    }
}
