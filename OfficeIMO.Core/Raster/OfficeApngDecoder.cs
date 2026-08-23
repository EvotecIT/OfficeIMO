using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Composes a selected, already-validated APNG frame onto its logical canvas.</summary>
internal static class OfficeApngDecoder {
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    internal static bool TryDecodeFrame(
        byte[] bytes,
        OfficeRasterContainerInfo container,
        int frameIndex,
        long maximumPixels,
        CancellationToken cancellationToken,
        out OfficeRasterImage? image) {
        image = null;
        try {
            if (container.Format != OfficeImageFormat.Png || !container.IsAnimated ||
                frameIndex < 0 || frameIndex >= container.Count ||
                !OfficeRasterImageDecoder.IsWithinPixelLimit(container.CanvasWidth, container.CanvasHeight, maximumPixels) ||
                !TryReadFrames(bytes, container, frameIndex, cancellationToken, out byte[] ihdr, out byte[]? palette,
                    out byte[]? transparency, out List<ApngFramePayload> framePayloads)) {
                return false;
            }

            for (int index = 0; index <= frameIndex; index++) {
                if (!IsFrameWithinMemoryBudget(bytes.Length, container, container.Frames[index], ihdr,
                        palette, transparency, framePayloads, index, index < frameIndex)) return false;
            }
            var canvas = new OfficeRasterImage(container.CanvasWidth, container.CanvasHeight, OfficeColor.Transparent);
            for (int index = 0; index <= frameIndex; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeRasterFrameInfo frame = container.Frames[index];
                byte[] standalone = CreateFramePng(ihdr, palette, transparency, frame.Width, frame.Height,
                    framePayloads[index], cancellationToken);
                if (!OfficePngReader.TryDecode(standalone, cancellationToken, out OfficeRasterImage? decoded) || decoded == null ||
                    decoded.Width != frame.Width || decoded.Height != frame.Height) {
                    return false;
                }

                byte[]? previous = index < frameIndex && frame.Disposal == OfficeRasterFrameDisposal.Previous
                    ? (byte[])canvas.PixelBuffer.Clone()
                    : null;
                Composite(canvas, decoded, frame, cancellationToken);
                if (index == frameIndex) {
                    image = canvas;
                    return true;
                }

                if (frame.Disposal == OfficeRasterFrameDisposal.Background) {
                    Clear(canvas, frame, cancellationToken);
                } else if (frame.Disposal == OfficeRasterFrameDisposal.Previous && previous != null) {
                    Buffer.BlockCopy(previous, 0, canvas.PixelBuffer, 0, previous.Length);
                }
            }
            return false;
        } catch (OperationCanceledException) {
            throw;
        } catch {
            image = null;
            return false;
        }
    }

    private static bool TryReadFrames(
        byte[] bytes,
        OfficeRasterContainerInfo container,
        int selectedFrameIndex,
        CancellationToken cancellationToken,
        out byte[] ihdr,
        out byte[]? palette,
        out byte[]? transparency,
        out List<ApngFramePayload> framePayloads) {
        ihdr = Array.Empty<byte>();
        palette = null;
        transparency = null;
        framePayloads = new List<ApngFramePayload>(selectedFrameIndex + 1);
        ApngFramePayload? current = null;
        int frameControlIndex = -1;
        bool seenImageData = false;
        bool currentUsesIdat = false;
        int segmentCount = 0;
        int cursor = Signature.Length;
        while (cursor <= bytes.Length - 12) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadInt32BigEndian(bytes, cursor);
            if (length < 0 || cursor > bytes.Length - 12 - length) return false;
            int dataOffset = cursor + 8;
            string type = Encoding.ASCII.GetString(bytes, cursor + 4, 4);
            if (type == "IHDR") {
                if (length != 13) return false;
                ihdr = new byte[length];
                Buffer.BlockCopy(bytes, dataOffset, ihdr, 0, length);
            } else if (type == "PLTE") {
                palette = Copy(bytes, dataOffset, length);
            } else if (type == "tRNS") {
                transparency = Copy(bytes, dataOffset, length);
            } else if (type == "fcTL") {
                if (current != null) {
                    framePayloads.Add(current);
                    if (frameControlIndex == selectedFrameIndex) {
                        return ihdr.Length == 13 && framePayloads.Count == selectedFrameIndex + 1;
                    }
                }
                frameControlIndex++;
                if (frameControlIndex > selectedFrameIndex) return false;
                current = new ApngFramePayload(bytes);
                currentUsesIdat = frameControlIndex == 0 && !seenImageData;
            } else if (type == "IDAT") {
                seenImageData = true;
                if (current != null && currentUsesIdat &&
                    !TryAddSegment(current, dataOffset, length, bytes.Length, container, ref segmentCount)) return false;
            } else if (type == "fdAT") {
                if (current == null || currentUsesIdat || length < 4) return false;
                if (!TryAddSegment(current, dataOffset + 4, length - 4, bytes.Length, container,
                        ref segmentCount)) return false;
            } else if (type == "IEND") {
                if (current != null) framePayloads.Add(current);
                return ihdr.Length == 13 && framePayloads.Count == selectedFrameIndex + 1;
            }
            cursor = checked(cursor + 12 + length);
        }
        return false;
    }

    private static bool TryAddSegment(
        ApngFramePayload payload,
        int offset,
        int length,
        int encodedLength,
        OfficeRasterContainerInfo container,
        ref int segmentCount) {
        long canvasBytes = checked((long)container.CanvasWidth * container.CanvasHeight * 4L);
        long structuralBytes = checked(encodedLength + canvasBytes + container.Count * 128L +
            ((long)segmentCount + 1L) * 32L);
        if (length < 0 || structuralBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
        payload.Add(offset, length);
        segmentCount++;
        return true;
    }

    private static byte[] CreateFramePng(
        byte[] sourceIhdr,
        byte[]? palette,
        byte[]? transparency,
        int width,
        int height,
        ApngFramePayload compressed,
        CancellationToken cancellationToken) {
        byte[] ihdr = (byte[])sourceIhdr.Clone();
        WriteInt32BigEndian(ihdr, 0, width);
        WriteInt32BigEndian(ihdr, 4, height);
        using var output = new MemoryStream(GetStandalonePngLength(compressed, palette, transparency));
        output.Write(Signature, 0, Signature.Length);
        WriteChunk(output, "IHDR", ihdr, cancellationToken);
        if (palette != null) WriteChunk(output, "PLTE", palette, cancellationToken);
        if (transparency != null) WriteChunk(output, "tRNS", transparency, cancellationToken);
        foreach (ApngFrameSegment segment in compressed.Segments) {
            WriteChunk(output, "IDAT", compressed.Source, segment.Offset, segment.Length, cancellationToken);
        }
        WriteChunk(output, "IEND", Array.Empty<byte>(), cancellationToken);
        return output.ToArray();
    }

    private static bool IsFrameWithinMemoryBudget(
        int encodedLength,
        OfficeRasterContainerInfo container,
        OfficeRasterFrameInfo frame,
        byte[] ihdr,
        byte[]? palette,
        byte[]? transparency,
        List<ApngFramePayload> payloads,
        int frameIndex,
        bool applyDisposal) {
        int standaloneLength = GetStandalonePngLength(payloads[frameIndex], palette, transparency);
        long canvasBytes = checked((long)container.CanvasWidth * container.CanvasHeight * 4L);
        long framePixels = checked((long)frame.Width * frame.Height);
        long frameRgbaBytes = checked(framePixels * 4L);
        long maximumScanlineBytes = checked(framePixels * 8L + frame.Height * 7L);
        long compressedBytes = payloads[frameIndex].CompressedLength;
        long segmentBytes = 0L;
        for (int index = 0; index < payloads.Count; index++) {
            segmentBytes = checked(segmentBytes + payloads[index].Segments.Count * 32L);
        }
        long peakBytes = checked(
            encodedLength + ihdr.Length + (palette?.Length ?? 0) + (transparency?.Length ?? 0) +
            container.Count * 128L + segmentBytes + canvasBytes +
            (applyDisposal && frame.Disposal == OfficeRasterFrameDisposal.Previous ? canvasBytes : 0L) +
            standaloneLength * 2L + compressedBytes * 3L + maximumScanlineBytes +
            frameRgbaBytes + frame.Width * 16L);
        return peakBytes <= OfficeRasterGuards.MaximumDecodedBytes;
    }

    private static int GetStandalonePngLength(
        ApngFramePayload payload,
        byte[]? palette,
        byte[]? transparency) {
        long length = Signature.Length + 25L + 12L;
        if (palette != null) length = checked(length + 12L + palette.Length);
        if (transparency != null) length = checked(length + 12L + transparency.Length);
        length = checked(length + payload.CompressedLength + payload.Segments.Count * 12L);
        if (length > int.MaxValue) throw new FormatException("APNG frame payload exceeds managed limits.");
        return (int)length;
    }

    private static void Composite(
        OfficeRasterImage canvas,
        OfficeRasterImage frameImage,
        OfficeRasterFrameInfo frame,
        CancellationToken cancellationToken) {
        byte[] source = frameImage.PixelBuffer;
        byte[] target = canvas.PixelBuffer;
        for (int y = 0; y < frame.Height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                int sourceOffset = ((y * frame.Width) + x) * 4;
                int targetOffset = (((frame.Y + y) * canvas.Width) + frame.X + x) * 4;
                if (frame.Blend == OfficeRasterFrameBlend.Source) {
                    Buffer.BlockCopy(source, sourceOffset, target, targetOffset, 4);
                } else {
                    canvas.BlendPixel(frame.X + x, frame.Y + y, OfficeColor.FromRgba(
                        source[sourceOffset], source[sourceOffset + 1], source[sourceOffset + 2], source[sourceOffset + 3]));
                }
            }
        }
    }

    private static void Clear(
        OfficeRasterImage canvas,
        OfficeRasterFrameInfo frame,
        CancellationToken cancellationToken) {
        byte[] pixels = canvas.PixelBuffer;
        for (int y = 0; y < frame.Height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            int offset = (((frame.Y + y) * canvas.Width) + frame.X) * 4;
            Array.Clear(pixels, offset, frame.Width * 4);
        }
    }

    private static void WriteChunk(
        Stream output,
        string type,
        byte[] data,
        CancellationToken cancellationToken) {
        WriteChunk(output, type, data, 0, data.Length, cancellationToken);
    }

    private static void WriteChunk(
        Stream output,
        string type,
        byte[] data,
        int offset,
        int length,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        byte[] header = new byte[8];
        WriteInt32BigEndian(header, 0, length);
        Buffer.BlockCopy(typeBytes, 0, header, 4, 4);
        output.Write(header, 0, header.Length);
        output.Write(data, offset, length);
        uint crc = ComputeCrc(typeBytes, data, offset, length, cancellationToken);
        byte[] checksum = new byte[4];
        WriteUInt32BigEndian(checksum, 0, crc);
        output.Write(checksum, 0, checksum.Length);
    }

    private static uint ComputeCrc(
        byte[] type,
        byte[] data,
        int offset,
        int length,
        CancellationToken cancellationToken) {
        uint crc = 0xFFFFFFFFU;
        for (int index = 0; index < type.Length; index++) crc = UpdateCrc(crc, type[index]);
        for (int index = 0; index < length; index++) {
            if ((index & 16383) == 0) cancellationToken.ThrowIfCancellationRequested();
            crc = UpdateCrc(crc, data[offset + index]);
        }
        return crc ^ 0xFFFFFFFFU;
    }

    private sealed class ApngFramePayload {
        private readonly List<ApngFrameSegment> _segments = new List<ApngFrameSegment>();

        internal ApngFramePayload(byte[] source) {
            Source = source;
        }

        internal byte[] Source { get; }
        internal IReadOnlyList<ApngFrameSegment> Segments => _segments;
        internal int CompressedLength { get; private set; }

        internal void Add(int offset, int length) {
            _segments.Add(new ApngFrameSegment(offset, length));
            CompressedLength = checked(CompressedLength + length);
        }
    }

    private readonly struct ApngFrameSegment {
        internal ApngFrameSegment(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }
        internal int Length { get; }
    }

    private static uint UpdateCrc(uint crc, byte value) {
        crc ^= value;
        for (int bit = 0; bit < 8; bit++) crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        return crc;
    }

    private static byte[] Copy(byte[] bytes, int offset, int length) {
        byte[] result = new byte[length];
        Buffer.BlockCopy(bytes, offset, result, 0, length);
        return result;
    }

    private static int ReadInt32BigEndian(byte[] bytes, int offset) =>
        bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];

    private static void WriteInt32BigEndian(byte[] bytes, int offset, int value) =>
        WriteUInt32BigEndian(bytes, offset, unchecked((uint)value));

    private static void WriteUInt32BigEndian(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }
}
