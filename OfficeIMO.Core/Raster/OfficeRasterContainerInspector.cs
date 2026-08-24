using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Reads bounded frame, timing, page, disposal, and blending information without decoding pixels.</summary>
public static class OfficeRasterContainerInspector {
    /// <summary>Attempts to inspect a bounded encoded raster container.</summary>
    public static bool TryInspect(byte[]? encodedBytes, out OfficeRasterContainerInfo? container) =>
        TryInspect(encodedBytes, options: null, out container);

    /// <summary>Attempts to inspect a bounded encoded raster container using per-request limits.</summary>
    public static bool TryInspect(
        byte[]? encodedBytes,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterContainerInfo? container) {
        container = null;
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (encodedBytes == null || encodedBytes.Length == 0 || encodedBytes.Length > effective.MaximumEncodedBytes ||
            !OfficeImageReader.TryIdentifyByContent(
                encodedBytes, fileName: null, effective.CancellationToken, out OfficeImageInfo imageInfo) ||
            imageInfo.Format != OfficeImageFormat.Tiff &&
            !OfficeRasterImageDecoder.IsWithinPixelLimit(imageInfo.Width, imageInfo.Height, effective.MaximumDecodedPixels)) {
            return false;
        }

        switch (imageInfo.Format) {
            case OfficeImageFormat.Gif:
                return TryInspectGif(encodedBytes, imageInfo, effective, out container);
            case OfficeImageFormat.Png:
                return TryInspectPng(encodedBytes, imageInfo, effective, out container);
            case OfficeImageFormat.Tiff:
                return OfficeTiffCodec.TryInspectPages(encodedBytes, effective, out container);
            case OfficeImageFormat.Webp:
                return TryInspectWebp(encodedBytes, imageInfo, effective, out container);
            default:
                container = CreateStatic(imageInfo);
                return true;
        }
    }

    /// <summary>Attempts to inspect a bounded encoded raster stream and leaves a seekable stream at its original position.</summary>
    public static bool TryInspect(
        Stream stream,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterContainerInfo? container) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        long originalPosition = stream.CanSeek ? stream.Position : 0L;
        try {
            if (!OfficeBoundedStreamReader.TryRead(stream, effective.MaximumEncodedBytes, effective.CancellationToken, out byte[] bytes)) {
                container = null;
                return false;
            }
            return TryInspect(bytes, effective, out container);
        } finally {
            if (stream.CanSeek) stream.Position = originalPosition;
        }
    }

    private static bool TryInspectGif(
        byte[] bytes,
        OfficeImageInfo imageInfo,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (bytes.Length < 14) return false;
        int cursor = 13;
        int packed = bytes[10];
        OfficeColor[]? globalColorTable = null;
        if ((packed & 0x80) != 0) {
            int colorCount = 1 << ((packed & 7) + 1);
            if (cursor > bytes.Length - colorCount * 3) return false;
            globalColorTable = new OfficeColor[colorCount];
            for (int index = 0; index < colorCount; index++) {
                int colorOffset = cursor + index * 3;
                globalColorTable[index] = OfficeColor.FromRgb(
                    bytes[colorOffset], bytes[colorOffset + 1], bytes[colorOffset + 2]);
            }
            cursor = checked(cursor + colorCount * 3);
        }
        int backgroundColorIndex = bytes[11];
        OfficeColor background = OfficeColor.Transparent;
        int loopCount = 1;
        bool hasLoopExtension = false;
        int delayHundredths = 0;
        int transparentIndex = -1;
        OfficeRasterFrameDisposal disposal = OfficeRasterFrameDisposal.None;
        var frames = new List<OfficeRasterFrameInfo>();
        bool sawTrailer = false;
        while (cursor < bytes.Length) {
            options.CancellationToken.ThrowIfCancellationRequested();
            byte introducer = bytes[cursor++];
            if (introducer == 0x3B) {
                if (cursor != bytes.Length) return false;
                sawTrailer = true;
                break;
            }
            if (introducer == 0x21) {
                if (cursor >= bytes.Length) return false;
                byte label = bytes[cursor++];
                if (label == 0xF9) {
                    if (cursor > bytes.Length - 6 || bytes[cursor] != 4 || bytes[cursor + 5] != 0) return false;
                    int control = bytes[cursor + 1];
                    int disposalCode = (control >> 2) & 7;
                    disposal = disposalCode switch {
                        2 => OfficeRasterFrameDisposal.Background,
                        3 => OfficeRasterFrameDisposal.Previous,
                        _ => OfficeRasterFrameDisposal.None
                    };
                    delayHundredths = bytes[cursor + 2] | bytes[cursor + 3] << 8;
                    transparentIndex = (control & 1) != 0 ? bytes[cursor + 4] : -1;
                    cursor += 6;
                } else if (label == 0xFF) {
                    if (cursor >= bytes.Length) return false;
                    int headerLength = bytes[cursor++];
                    if (headerLength < 1 || cursor > bytes.Length - headerLength) return false;
                    bool netscape = headerLength == 11 &&
                        HasAscii(bytes, cursor, "NETSCAPE2.0");
                    cursor += headerLength;
                    if (netscape && cursor <= bytes.Length - 5 && bytes[cursor] == 3 && bytes[cursor + 1] == 1) {
                        loopCount = bytes[cursor + 2] | bytes[cursor + 3] << 8;
                        hasLoopExtension = true;
                    }
                    if (!SkipSubBlocks(bytes, ref cursor, options.CancellationToken)) return false;
                } else if (label == 0x01) {
                    if (!SkipSubBlocks(bytes, ref cursor, options.CancellationToken)) return false;
                    delayHundredths = 0;
                    disposal = OfficeRasterFrameDisposal.None;
                    transparentIndex = -1;
                } else if (!SkipSubBlocks(bytes, ref cursor, options.CancellationToken)) {
                    return false;
                }
                continue;
            }
            if (introducer != 0x2C || cursor > bytes.Length - 9) return false;
            int x = ReadUInt16LittleEndian(bytes, cursor);
            int y = ReadUInt16LittleEndian(bytes, cursor + 2);
            int width = ReadUInt16LittleEndian(bytes, cursor + 4);
            int height = ReadUInt16LittleEndian(bytes, cursor + 6);
            int descriptorPacked = bytes[cursor + 8];
            cursor += 9;
            if (width < 1 || height < 1 || x > imageInfo.Width - width || y > imageInfo.Height - height) return false;
            if ((descriptorPacked & 0x80) != 0) {
                int localTableBytes = 3 << ((descriptorPacked & 7) + 1);
                if (cursor > bytes.Length - localTableBytes) return false;
                cursor += localTableBytes;
            }
            if (cursor >= bytes.Length) return false;
            int minimumCodeSize = bytes[cursor++];
            if (minimumCodeSize < 2 || minimumCodeSize > 8) return false;
            if (!SkipSubBlocks(bytes, ref cursor, options.CancellationToken)) return false;
            if (frames.Count >= 65535) return false;
            if (frames.Count == 0 && globalColorTable != null &&
                backgroundColorIndex >= 0 && backgroundColorIndex < globalColorTable.Length &&
                backgroundColorIndex != transparentIndex) {
                background = globalColorTable[backgroundColorIndex];
            }
            frames.Add(new OfficeRasterFrameInfo(
                frames.Count,
                OfficeRasterFrameKind.AnimationFrame,
                width,
                height,
                x,
                y,
                TimeSpan.FromMilliseconds(delayHundredths * 10D),
                disposal,
                OfficeRasterFrameBlend.Over,
                frames.Count == 0));
            delayHundredths = 0;
            disposal = OfficeRasterFrameDisposal.None;
            transparentIndex = -1;
        }
        if (!sawTrailer || frames.Count == 0) return false;
        if (frames.Count == 1 && frames[0].Duration == TimeSpan.Zero && !hasLoopExtension) {
            OfficeRasterFrameInfo frame = frames[0];
            frames[0] = new OfficeRasterFrameInfo(
                frame.Index,
                OfficeRasterFrameKind.Image,
                frame.Width,
                frame.Height,
                frame.X,
                frame.Y,
                frame.Duration,
                frame.Disposal,
                frame.Blend,
                frame.IsDefaultImage);
        }
        container = new OfficeRasterContainerInfo(
            OfficeImageFormat.Gif,
            imageInfo.Width,
            imageInfo.Height,
            frames.ToArray(),
            loopCount,
            background);
        return true;
    }

    private static bool TryInspectPng(
        byte[] bytes,
        OfficeImageInfo imageInfo,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (!OfficePngReader.TryGetFrameCount(
                bytes, options.CancellationToken, out int frameCount)) return false;
        if (frameCount > 65535) return false;
        int loopCount = 0;
        bool seenImageData = false;
        bool seenAnimationControl = false;
        var frames = new List<OfficeRasterFrameInfo>(frameCount);
        int cursor = 8;
        while (cursor <= bytes.Length - 12) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int length = ReadInt32BigEndian(bytes, cursor);
            if (length < 0 || cursor > bytes.Length - 12 - length) return false;
            int data = cursor + 8;
            string type = ReadAscii(bytes, cursor + 4);
            if (type == "acTL") {
                if (length != 8) return false;
                seenAnimationControl = true;
                uint encodedLoopCount = ReadUInt32BigEndian(bytes, data + 4);
                if (encodedLoopCount > int.MaxValue) return false;
                loopCount = (int)encodedLoopCount;
            } else if (type == "fcTL") {
                if (length != 26) return false;
                int width = ReadInt32BigEndian(bytes, data + 4);
                int height = ReadInt32BigEndian(bytes, data + 8);
                int x = ReadInt32BigEndian(bytes, data + 12);
                int y = ReadInt32BigEndian(bytes, data + 16);
                int delayNumerator = ReadUInt16BigEndian(bytes, data + 20);
                int delayDenominator = ReadUInt16BigEndian(bytes, data + 22);
                if (delayDenominator == 0) delayDenominator = 100;
                int disposalCode = bytes[data + 24];
                int blendCode = bytes[data + 25];
                if (width < 1 || height < 1 || x < 0 || y < 0 ||
                    x > imageInfo.Width - width || y > imageInfo.Height - height ||
                    disposalCode > 2 || blendCode > 1) return false;
                frames.Add(new OfficeRasterFrameInfo(
                    frames.Count,
                    OfficeRasterFrameKind.AnimationFrame,
                    width,
                    height,
                    x,
                    y,
                    TimeSpan.FromSeconds(delayNumerator / (double)delayDenominator),
                    disposalCode switch {
                        1 => OfficeRasterFrameDisposal.Background,
                        2 when frames.Count == 0 => OfficeRasterFrameDisposal.Background,
                        2 => OfficeRasterFrameDisposal.Previous,
                        _ => OfficeRasterFrameDisposal.None
                    },
                    blendCode == 0 ? OfficeRasterFrameBlend.Source : OfficeRasterFrameBlend.Over,
                    frames.Count == 0 && !seenImageData));
            } else if (type == "IDAT") {
                seenImageData = true;
            }
            cursor = checked(cursor + 12 + length);
            if (type == "IEND") break;
        }
        if (!seenAnimationControl) {
            container = CreateStatic(imageInfo);
            return true;
        }
        if (frames.Count != frameCount ||
            !OfficePngAnimationValidator.TryValidateStructure(bytes, options.CancellationToken)) return false;
        container = new OfficeRasterContainerInfo(
            OfficeImageFormat.Png,
            imageInfo.Width,
            imageInfo.Height,
            frames.ToArray(),
            loopCount,
            OfficeColor.Transparent);
        return true;
    }

    private static bool TryInspectWebp(
        byte[] bytes,
        OfficeImageInfo imageInfo,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) {
        container = null;
        var frames = new List<OfficeRasterFrameInfo>();
        int loopCount = 1;
        OfficeColor background = OfficeColor.Transparent;
        int cursor = 12;
        while (cursor <= bytes.Length - 8) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int length = ReadInt32LittleEndian(bytes, cursor + 4);
            if (length < 0 || cursor > bytes.Length - 8 - length) return false;
            int data = cursor + 8;
            string type = ReadAscii(bytes, cursor);
            if (type == "ANIM") {
                if (length != 6) return false;
                background = new OfficeColor(bytes[data + 2], bytes[data + 1], bytes[data], bytes[data + 3]);
                loopCount = ReadUInt16LittleEndian(bytes, data + 4);
            } else if (type == "ANMF") {
                if (length < 16) return false;
                if (frames.Count >= 65535) return false;
                int x = checked(ReadUInt24LittleEndian(bytes, data) * 2);
                int y = checked(ReadUInt24LittleEndian(bytes, data + 3) * 2);
                int width = checked(ReadUInt24LittleEndian(bytes, data + 6) + 1);
                int height = checked(ReadUInt24LittleEndian(bytes, data + 9) + 1);
                int durationMs = ReadUInt24LittleEndian(bytes, data + 12);
                int flags = bytes[data + 15];
                if ((flags & 0xFC) != 0 || width < 1 || height < 1 ||
                    x > imageInfo.Width - width || y > imageInfo.Height - height) return false;
                frames.Add(new OfficeRasterFrameInfo(
                    frames.Count,
                    OfficeRasterFrameKind.AnimationFrame,
                    width,
                    height,
                    x,
                    y,
                    TimeSpan.FromMilliseconds(durationMs),
                    (flags & 1) == 0 ? OfficeRasterFrameDisposal.None : OfficeRasterFrameDisposal.Background,
                    (flags & 2) == 0 ? OfficeRasterFrameBlend.Over : OfficeRasterFrameBlend.Source,
                    isDefaultImage: false));
            }
            cursor = checked(data + length + (length & 1));
        }
        if (cursor != bytes.Length) return false;
        if (frames.Count == 0) {
            container = CreateStatic(imageInfo);
            return true;
        }
        container = new OfficeRasterContainerInfo(
            OfficeImageFormat.Webp,
            imageInfo.Width,
            imageInfo.Height,
            frames.ToArray(),
            loopCount,
            background);
        return true;
    }

    private static OfficeRasterContainerInfo CreateStatic(OfficeImageInfo info) =>
        new OfficeRasterContainerInfo(
            info.Format,
            info.Width,
            info.Height,
            new[] {
                new OfficeRasterFrameInfo(
                    0,
                    OfficeRasterFrameKind.Image,
                    info.Width,
                    info.Height,
                    0,
                    0,
                    TimeSpan.Zero,
                    OfficeRasterFrameDisposal.None,
                    OfficeRasterFrameBlend.Source,
                    true)
            },
            1,
            OfficeColor.Transparent);

    private static bool SkipSubBlocks(
        byte[] bytes,
        ref int cursor,
        System.Threading.CancellationToken cancellationToken) {
        int blockCount = 0;
        while (cursor < bytes.Length) {
            if ((blockCount++ & 0x3FF) == 0) cancellationToken.ThrowIfCancellationRequested();
            int length = bytes[cursor++];
            if (length == 0) return true;
            if (cursor > bytes.Length - length) return false;
            cursor += length;
        }
        return false;
    }

    private static bool HasAscii(byte[] bytes, int offset, string text) {
        if (offset < 0 || offset > bytes.Length - text.Length) return false;
        for (int index = 0; index < text.Length; index++) {
            if (bytes[offset + index] != (byte)text[index]) return false;
        }
        return true;
    }

    private static string ReadAscii(byte[] bytes, int offset) =>
        new string(new[] { (char)bytes[offset], (char)bytes[offset + 1], (char)bytes[offset + 2], (char)bytes[offset + 3] });

    private static int ReadInt32BigEndian(byte[] bytes, int offset) =>
        bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];

    private static uint ReadUInt32BigEndian(byte[] bytes, int offset) =>
        (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16 |
        (uint)bytes[offset + 2] << 8 | bytes[offset + 3];

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;

    private static int ReadUInt16BigEndian(byte[] bytes, int offset) =>
        bytes[offset] << 8 | bytes[offset + 1];

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8;

    private static int ReadUInt24LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16;
}
