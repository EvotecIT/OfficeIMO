using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Reads bounded frame, timing, page, disposal, and blending information after validating managed payloads.</summary>
public static class OfficeRasterContainerInspector {
    /// <summary>Attempts to inspect a bounded encoded raster container.</summary>
    public static bool TryInspect(byte[]? encodedBytes, out OfficeRasterContainerInfo? container) =>
        TryInspect(encodedBytes, options: null, out container);

    /// <summary>Attempts to inspect a bounded encoded raster container using per-request limits.</summary>
    public static bool TryInspect(
        byte[]? encodedBytes,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterContainerInfo? container) =>
        TryInspectCore(
            encodedBytes, options, enforceAllTiffPagePixelLimits: true, out container, out _);

    internal static bool TryInspectForDecode(
        byte[]? encodedBytes,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container,
        out OfficeImageFormat detectedFormat) =>
        TryInspectCore(
            encodedBytes, options, enforceAllTiffPagePixelLimits: false,
            out container, out detectedFormat);

    private static bool TryInspectCore(
        byte[]? encodedBytes,
        OfficeRasterDecodeOptions? options,
        bool enforceAllTiffPagePixelLimits,
        out OfficeRasterContainerInfo? container,
        out OfficeImageFormat detectedFormat) {
        container = null;
        detectedFormat = OfficeImageFormat.Unknown;
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (encodedBytes == null || encodedBytes.Length == 0 || encodedBytes.Length > effective.MaximumEncodedBytes ||
            !IsInspectionWorkingSetWithinLimit(encodedBytes.LongLength, effective.RetainedManagedBytes, frameCount: 0) ||
            !OfficeImageReader.TryIdentifyByContent(
                encodedBytes, fileName: null, effective.CancellationToken, out OfficeImageInfo imageInfo)) {
            return false;
        }
        detectedFormat = imageInfo.Format;
        if (imageInfo.Format != OfficeImageFormat.Tiff &&
            !OfficeRasterImageDecoder.IsWithinPixelLimit(imageInfo.Width, imageInfo.Height, effective.MaximumDecodedPixels)) {
            return false;
        }

        switch (imageInfo.Format) {
            case OfficeImageFormat.Gif:
                return TryInspectGif(encodedBytes, imageInfo, effective, out container);
            case OfficeImageFormat.Png:
                return TryInspectPng(
                    encodedBytes, imageInfo, effective,
                    validateDecodedPayloads: enforceAllTiffPagePixelLimits, out container);
            case OfficeImageFormat.Tiff:
                return OfficeTiffCodec.TryInspectPages(
                    encodedBytes, effective, enforceAllTiffPagePixelLimits, out container);
            case OfficeImageFormat.Webp:
                return TryInspectWebp(encodedBytes, imageInfo, effective, out container);
            case OfficeImageFormat.Jpeg:
                return TryInspectJpeg(
                    encodedBytes, imageInfo, effective,
                    validateDecodedPayload: enforceAllTiffPagePixelLimits, out container);
            case OfficeImageFormat.Bmp:
                if (!OfficeBmpReader.TryValidatePayload(encodedBytes, effective.CancellationToken)) return false;
                container = CreateStatic(imageInfo);
                return true;
            default:
                return false;
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
            if (!OfficeBoundedStreamReader.TryRead(
                    stream, effective.MaximumEncodedBytes, effective.CancellationToken,
                    out byte[] bytes, out long retainedManagedBytes)) {
                container = null;
                return false;
            }
            return TryInspect(
                bytes,
                effective.WithAdditionalRetainedManagedBytes(retainedManagedBytes),
                out container);
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
        if (globalColorTable == null
                ? backgroundColorIndex != 0
                : backgroundColorIndex >= globalColorTable.Length) return false;
        OfficeColor background = OfficeColor.Transparent;
        int loopCount = 1;
        bool hasLoopExtension = false;
        int delayHundredths = 0;
        int transparentIndex = -1;
        OfficeRasterFrameDisposal disposal = OfficeRasterFrameDisposal.None;
        bool hasPendingGraphicControl = false;
        var frames = new List<OfficeRasterFrameInfo>();
        bool sawTrailer = false;
        long decodedFramePixels = 0L;
        while (cursor < bytes.Length) {
            options.CancellationToken.ThrowIfCancellationRequested();
            byte introducer = bytes[cursor++];
            if (introducer == 0x3B) {
                if (cursor != bytes.Length) return false;
                sawTrailer = true;
                break;
            }
            if (introducer == 0x21) {
                if (bytes[3] == (byte)'8' && bytes[4] == (byte)'7') return false;
                if (cursor >= bytes.Length) return false;
                byte label = bytes[cursor++];
                if (label == 0xF9) {
                    if (cursor > bytes.Length - 6 || bytes[cursor] != 4 || bytes[cursor + 5] != 0) return false;
                    int control = bytes[cursor + 1];
                    int disposalCode = (control >> 2) & 7;
                    if ((control & 0xE0) != 0 || disposalCode > 3 || hasPendingGraphicControl) return false;
                    hasPendingGraphicControl = true;
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
                    if (headerLength != 11 || cursor > bytes.Length - headerLength) return false;
                    bool netscape = headerLength == 11 &&
                        HasAscii(bytes, cursor, "NETSCAPE2.0");
                    cursor += headerLength;
                    if (netscape && cursor <= bytes.Length - 5 && bytes[cursor] == 3 && bytes[cursor + 1] == 1) {
                        loopCount = bytes[cursor + 2] | bytes[cursor + 3] << 8;
                        hasLoopExtension = true;
                    }
                    if (!SkipSubBlocks(bytes, ref cursor, options.CancellationToken)) return false;
                } else if (label == 0x01) {
                    // Plain Text Extensions are graphic rendering blocks. Until the managed GIF
                    // decoder can render and inventory them, fail closed rather than lose content.
                    return false;
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
            if ((descriptorPacked & 0x18) != 0 || width < 1 || height < 1 ||
                x > imageInfo.Width - width || y > imageInfo.Height - height) return false;
            int activeColorCount = globalColorTable?.Length ?? 0;
            if ((descriptorPacked & 0x80) != 0) {
                activeColorCount = 1 << ((descriptorPacked & 7) + 1);
                int localTableBytes = activeColorCount * 3;
                if (cursor > bytes.Length - localTableBytes) return false;
                cursor += localTableBytes;
            }
            if (activeColorCount == 0 || transparentIndex >= activeColorCount) return false;
            if (cursor >= bytes.Length) return false;
            int minimumCodeSize = bytes[cursor++];
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out int framePixels) ||
                decodedFramePixels > OfficeRasterGuards.MaximumPixels - framePixels) return false;
            decodedFramePixels += framePixels;
            if (!OfficeGifReader.TryValidateImageData(
                    bytes,
                    ref cursor,
                    minimumCodeSize,
                    framePixels,
                    activeColorCount,
                    rejectTrailingLzwBytes: false,
                    options.CancellationToken,
                    checked(options.RetainedManagedBytes + frames.Count * 128L))) return false;
            if (frames.Count >= 65535 ||
                !IsInspectionWorkingSetWithinLimit(
                    bytes.LongLength, options.RetainedManagedBytes, frames.Count + 1)) return false;
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
            hasPendingGraphicControl = false;
            delayHundredths = 0;
            disposal = OfficeRasterFrameDisposal.None;
            transparentIndex = -1;
        }
        if (!sawTrailer || frames.Count == 0 || hasPendingGraphicControl) return false;
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

    private static bool TryInspectJpeg(
        byte[] bytes,
        OfficeImageInfo imageInfo,
        OfficeRasterDecodeOptions options,
        bool validateDecodedPayload,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (!OfficeImageReader.HasCompleteJpegPayload(
                bytes,
                options.CancellationToken,
                requireManagedFrame: true,
                validateMetadata: false)) return false;
        if (validateDecodedPayload && (!OfficeJpegCodec.TryDecode(
                bytes,
                options.CancellationToken,
                checked(options.RetainedManagedBytes + 128L),
                out OfficeRasterImage? decoded) || decoded == null)) return false;
        int canvasWidth = imageInfo.Width;
        int canvasHeight = imageInfo.Height;
        if (OfficeImageOrientationNormalizer.TryRead(
                bytes, options.CancellationToken, out OfficeImageOrientation orientation) &&
            orientation is >= OfficeImageOrientation.Transpose and <= OfficeImageOrientation.Rotate90CounterClockwise) {
            canvasWidth = imageInfo.Height;
            canvasHeight = imageInfo.Width;
        }
        container = CreateStatic(imageInfo, canvasWidth, canvasHeight);
        return true;
    }

    private static bool TryInspectPng(
        byte[] bytes,
        OfficeImageInfo imageInfo,
        OfficeRasterDecodeOptions options,
        bool validateDecodedPayloads,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (!OfficePngReader.TryGetFrameCount(
                bytes, options.CancellationToken, out int frameCount)) return false;
        if (frameCount > 65535 ||
            !IsInspectionWorkingSetWithinLimit(
                bytes.LongLength, options.RetainedManagedBytes, frameCount)) return false;
        if (validateDecodedPayloads && !OfficePngReader.TryValidateDecodedPayload(
                bytes,
                options.CancellationToken,
                checked(options.RetainedManagedBytes + frameCount * 128L))) return false;
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
        bool hasLosslessImage = false;
        long validatedFramePixels = 0L;
        int cursor = 12;
        while (cursor <= bytes.Length - 8) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int length = ReadInt32LittleEndian(bytes, cursor + 4);
            if (length < 0 || cursor > bytes.Length - 8 - length) return false;
            int data = cursor + 8;
            string type = ReadAscii(bytes, cursor);
            if (type == "VP8L") {
                hasLosslessImage = true;
            } else if (type == "ANIM") {
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
                long framePixels = checked((long)width * height);
                if (framePixels > options.MaximumDecodedPixels - validatedFramePixels ||
                    !TryValidateWebpAnimationFramePayload(
                        bytes, data + 16, length - 16, width, height,
                        checked(options.RetainedManagedBytes + (frames.Count + 1L) * 128L),
                        options.CancellationToken)) return false;
                validatedFramePixels += framePixels;
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
            if (!hasLosslessImage ||
                !OfficeWebpCodec.TryDecode(
                    bytes, options.CancellationToken, options.RetainedManagedBytes,
                    out OfficeRasterImage? decoded) ||
                decoded == null || decoded.Width != imageInfo.Width || decoded.Height != imageInfo.Height) return false;
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

    private static bool TryValidateWebpAnimationFramePayload(
        byte[] source,
        int offset,
        int length,
        int expectedWidth,
        int expectedHeight,
        long retainedFrameInventoryBytes,
        System.Threading.CancellationToken cancellationToken) {
        if (length < 8 || offset < 0 || offset > source.Length - length) return false;
        int end = checked(offset + length);
        int cursor = offset;
        bool seenAlpha = false;
        bool seenImage = false;
        while (cursor < end) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cursor > end - 8) return false;
            int payloadLength = ReadInt32LittleEndian(source, cursor + 4);
            if (payloadLength < 0) return false;
            int payloadOffset = checked(cursor + 8);
            long payloadEnd = (long)payloadOffset + payloadLength;
            long paddedEnd = payloadEnd + (payloadLength & 1);
            if (payloadEnd > end || paddedEnd > end ||
                (payloadLength & 1) != 0 && source[(int)payloadEnd] != 0) return false;

            string type = ReadAscii(source, cursor);
            if (type == "ALPH") {
                if (seenAlpha || seenImage ||
                    !OfficeImageReader.HasValidWebpAlphaHeader(source, payloadOffset, payloadLength)) return false;
                seenAlpha = true;
            } else if (type == "VP8 " || type == "VP8L") {
                if (seenImage || seenAlpha && type == "VP8L" ||
                    !OfficeImageReader.TryReadWebpImageHeader(
                        source, payloadOffset, payloadLength, type,
                        out int width, out int height, out _) ||
                    width != expectedWidth || height != expectedHeight) return false;
                if (type == "VP8L" && !TryDecodeWebpAnimationFrame(
                        source, cursor, (int)paddedEnd - cursor,
                        expectedWidth, expectedHeight, retainedFrameInventoryBytes,
                        cancellationToken)) return false;
                seenImage = true;
            } else {
                return false;
            }
            cursor = (int)paddedEnd;
        }
        return cursor == end && seenImage;
    }

    private static bool IsInspectionWorkingSetWithinLimit(
        long encodedBytes,
        long retainedManagedBytes,
        int frameCount) {
        if (encodedBytes < 0L || retainedManagedBytes < 0L || frameCount < 0) return false;
        try {
            return checked(
                encodedBytes + retainedManagedBytes + frameCount * 128L + 64L * 1024L) <=
                OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryDecodeWebpAnimationFrame(
        byte[] source,
        int chunkOffset,
        int chunkLength,
        int expectedWidth,
        int expectedHeight,
        long retainedFrameInventoryBytes,
        System.Threading.CancellationToken cancellationToken) {
        int wrappedLength = checked(12 + chunkLength);
        long retainedManagedBytes = checked(
            source.LongLength + 24L + retainedFrameInventoryBytes + 64L * 1024L);
        if (retainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes - wrappedLength - 24L) return false;
        var wrapped = new byte[wrappedLength];
        wrapped[0] = (byte)'R';
        wrapped[1] = (byte)'I';
        wrapped[2] = (byte)'F';
        wrapped[3] = (byte)'F';
        int riffLength = wrappedLength - 8;
        wrapped[4] = (byte)riffLength;
        wrapped[5] = (byte)(riffLength >> 8);
        wrapped[6] = (byte)(riffLength >> 16);
        wrapped[7] = (byte)(riffLength >> 24);
        wrapped[8] = (byte)'W';
        wrapped[9] = (byte)'E';
        wrapped[10] = (byte)'B';
        wrapped[11] = (byte)'P';
        Buffer.BlockCopy(source, chunkOffset, wrapped, 12, chunkLength);
        return OfficeWebpCodec.TryDecode(
                   wrapped, cancellationToken, retainedManagedBytes, out OfficeRasterImage? decoded) &&
               decoded != null && decoded.Width == expectedWidth && decoded.Height == expectedHeight;
    }

    private static OfficeRasterContainerInfo CreateStatic(OfficeImageInfo info) =>
        CreateStatic(info, info.Width, info.Height);

    private static OfficeRasterContainerInfo CreateStatic(
        OfficeImageInfo info,
        int canvasWidth,
        int canvasHeight) =>
        new OfficeRasterContainerInfo(
            info.Format,
            canvasWidth,
            canvasHeight,
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
