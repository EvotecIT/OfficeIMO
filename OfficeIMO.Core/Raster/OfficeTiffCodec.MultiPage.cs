using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    /// <summary>Encodes one or more RGBA pages as a bounded classic TIFF IFD chain.</summary>
    public static byte[] EncodePages(
        IReadOnlyList<OfficeRasterImage> pages,
        OfficeTiffEncodeOptions? options = null) {
        if (pages == null) throw new ArgumentNullException(nameof(pages));
        if (pages.Count < 1 || pages.Count > 1024) throw new ArgumentOutOfRangeException(nameof(pages));
        OfficeTiffEncodeOptions effective = options ?? new OfficeTiffEncodeOptions();
        ValidateOptions(effective);

        long totalPixels = 0;
        for (int index = 0; index < pages.Count; index++) {
            OfficeRasterImage page = pages[index] ?? throw new ArgumentException("TIFF pages cannot contain null images.", nameof(pages));
            totalPixels = checked(totalPixels + (long)page.Width * page.Height);
            if (totalPixels > OfficeRasterGuards.MaximumPixels) {
                throw new ArgumentException("TIFF pages exceed the aggregate decoded-pixel limit.", nameof(pages));
            }
        }

        int entryCount = BaseEntryCount - (effective.WriteResolution ? 0 : 3) + (UsesHorizontalPredictor(effective) ? 1 : 0);
        int ifdBlockLength = 2 + entryCount * 12 + 4 + 8 + (effective.WriteResolution ? 16 : 0);
        long headerLength = checked(8L + (long)pages.Count * ifdBlockLength);
        long sourceBytes = checked(totalPixels * 4L);
        var strips = new byte[pages.Count][];
        long stripBytes = 0;
        long retainedStripBytes = 0;
        for (int index = 0; index < pages.Count; index++) {
            OfficeRasterImage page = pages[index];
            long pendingAllocationBytes = EstimateMultiPageStripEncodingPeak(page, effective);
            if (!CanBeginMultiPageStripEncoding(sourceBytes, retainedStripBytes, pendingAllocationBytes)) {
                throw new ArgumentException("The multi-page TIFF encoding working set exceeds the managed limit.", nameof(pages));
            }
            byte[] strip = EncodeTiffStrip(pages[index], effective);
            strips[index] = strip;
            stripBytes = checked(stripBytes + strip.Length);
            if (headerLength + stripBytes > OfficeRasterGuards.MaximumEncodedBytes) {
                throw new ArgumentException("The multi-page TIFF exceeds the encoded-size limit.", nameof(pages));
            }
            if (!ReferenceEquals(strip, pages[index].PixelBuffer)) {
                retainedStripBytes = checked(retainedStripBytes + strip.Length);
            }
            if (!IsMultiPageTiffWorkingSetWithinLimit(
                    sourceBytes,
                    retainedStripBytes,
                    checked(headerLength + stripBytes))) {
                throw new ArgumentException("The multi-page TIFF encoding working set exceeds the managed limit.", nameof(pages));
            }
        }

        int fileLength = OfficeRasterGuards.EnsureOutputBytes(
            checked(headerLength + stripBytes),
            "The multi-page TIFF exceeds the encoded-size limit.");
        var output = new byte[fileLength];
        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, 8);

        int stripOffset = checked((int)headerLength);
        for (int index = 0; index < pages.Count; index++) {
            OfficeRasterImage page = pages[index];
            int ifdOffset = checked(8 + index * ifdBlockLength);
            int bitsOffset = checked(ifdOffset + 2 + entryCount * 12 + 4);
            int xResolutionOffset = bitsOffset + 8;
            int yResolutionOffset = xResolutionOffset + 8;
            int nextIfdOffset = index + 1 < pages.Count ? ifdOffset + ifdBlockLength : 0;
            WritePageIfd(output, ifdOffset, page, effective, stripOffset, strips[index].Length,
                bitsOffset, xResolutionOffset, yResolutionOffset, nextIfdOffset);
            Buffer.BlockCopy(strips[index], 0, output, stripOffset, strips[index].Length);
            stripOffset = checked(stripOffset + strips[index].Length);
        }
        return output;
    }

    internal static bool IsMultiPageTiffWorkingSetWithinLimit(
        long sourceBytes,
        long retainedStripBytes,
        long outputBytes) {
        if (sourceBytes < 0L || retainedStripBytes < 0L || outputBytes < 0L) return false;
        try {
            return checked(sourceBytes + retainedStripBytes + outputBytes + 64L * 1024L) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool CanBeginMultiPageStripEncoding(
        long sourceBytes,
        long retainedStripBytes,
        long pendingAllocationBytes) {
        if (sourceBytes < 0L || retainedStripBytes < 0L || pendingAllocationBytes < 0L) return false;
        try {
            return checked(sourceBytes + retainedStripBytes + pendingAllocationBytes + 64L * 1024L) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool IsSinglePageTiffCompressionWorkingSetWithinLimit(
        long sourceBytes,
        OfficeTiffEncodeOptions options) {
        if (options == null || sourceBytes < 0L) return false;
        try {
            long pendingAllocationBytes = options.Compression switch {
                OfficeTiffCompression.None => 0L,
                OfficeTiffCompression.PackBits => 0L,
                OfficeTiffCompression.Lzw => checked(sourceBytes *
                    (UsesHorizontalPredictor(options) ? 6L : 5L)),
                OfficeTiffCompression.Deflate => checked(sourceBytes *
                    (UsesHorizontalPredictor(options) ? 5L : 4L)),
                _ => throw new ArgumentOutOfRangeException(nameof(options.Compression))
            };
            return CanBeginMultiPageStripEncoding(sourceBytes, 0L, pendingAllocationBytes);
        } catch (OverflowException) {
            return false;
        }
    }

    private static void EnsureSinglePageCompressionWorkingSet(
        long sourceBytes,
        OfficeTiffEncodeOptions options) {
        if (!IsSinglePageTiffCompressionWorkingSetWithinLimit(sourceBytes, options)) {
            throw new ArgumentException("The TIFF compression working set exceeds the managed limit.", nameof(options));
        }
    }

    private static long EstimateMultiPageStripEncodingPeak(
        OfficeRasterImage page,
        OfficeTiffEncodeOptions options) {
        long sourceLength = page.PixelBuffer.LongLength;
        return options.Compression switch {
            OfficeTiffCompression.None => 0L,
            OfficeTiffCompression.PackBits => EncodePackBitsRows(
                page.PixelBuffer, checked(page.Width * 4), page.Height, output: null, outputOffset: 0),
            OfficeTiffCompression.Lzw => checked(sourceLength *
                (UsesHorizontalPredictor(options) ? 6L : 5L)),
            OfficeTiffCompression.Deflate => checked(sourceLength *
                (UsesHorizontalPredictor(options) ? 5L : 4L)),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Compression))
        };
    }

    /// <summary>Encodes one or more RGBA pages to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodePagesTo(
        IReadOnlyList<OfficeRasterImage> pages,
        Stream destination,
        OfficeTiffEncodeOptions? options = null) {
        OfficeRasterOutput.EnsureWritable(destination);
        byte[] encoded = EncodePages(pages, options);
        destination.Write(encoded, 0, encoded.Length);
    }

    private static byte[] EncodeTiffStrip(OfficeRasterImage image, OfficeTiffEncodeOptions options) {
        byte[] pixels = image.PixelBuffer;
        switch (options.Compression) {
            case OfficeTiffCompression.None:
                return pixels;
            case OfficeTiffCompression.Lzw:
                return EncodeTiffLzw(pixels, image.Width, image.Height, options);
            case OfficeTiffCompression.PackBits:
                int length = EncodePackBitsRows(pixels, image.Width * 4, image.Height, null, 0);
                var packed = new byte[length];
                if (EncodePackBitsRows(pixels, image.Width * 4, image.Height, packed, 0) != length) {
                    throw new InvalidOperationException("TIFF PackBits length changed while encoding.");
                }
                return packed;
            case OfficeTiffCompression.Deflate:
                return OfficeZlibCodec.Compress(
                    PrepareTiffCompressionInput(pixels, image.Width, image.Height, options));
            default:
                throw new ArgumentOutOfRangeException(nameof(options.Compression));
        }
    }

    private static void WritePageIfd(
        byte[] output,
        int ifdOffset,
        OfficeRasterImage image,
        OfficeTiffEncodeOptions options,
        int stripOffset,
        int stripLength,
        int bitsOffset,
        int xResolutionOffset,
        int yResolutionOffset,
        int nextIfdOffset) {
        bool writePredictor = UsesHorizontalPredictor(options);
        int entryCount = BaseEntryCount - (options.WriteResolution ? 0 : 3) + (writePredictor ? 1 : 0);
        WriteUInt16(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, image.Width);
        WriteEntry(output, ref entry, 257, 4, 1, image.Height);
        WriteEntry(output, ref entry, 258, 3, 4, bitsOffset);
        WriteShortEntry(output, ref entry, 259, (int)options.Compression);
        WriteShortEntry(output, ref entry, 262, 2);
        WriteEntry(output, ref entry, 273, 4, 1, stripOffset);
        WriteShortEntry(output, ref entry, 274, 1);
        WriteShortEntry(output, ref entry, 277, 4);
        WriteEntry(output, ref entry, 278, 4, 1, image.Height);
        WriteEntry(output, ref entry, 279, 4, 1, stripLength);
        if (options.WriteResolution) {
            WriteEntry(output, ref entry, 282, 5, 1, xResolutionOffset);
            WriteEntry(output, ref entry, 283, 5, 1, yResolutionOffset);
        }
        WriteShortEntry(output, ref entry, 284, 1);
        if (options.WriteResolution) WriteShortEntry(output, ref entry, 296, 2);
        if (writePredictor) WriteShortEntry(output, ref entry, 317, (int)options.Predictor);
        WriteShortEntry(output, ref entry, 338, 2);
        WriteUInt32(output, entry, nextIfdOffset);
        for (int sample = 0; sample < 4; sample++) WriteUInt16(output, bitsOffset + sample * 2, 8);
        if (options.WriteResolution) {
            WriteRational(output, xResolutionOffset, options.DpiX);
            WriteRational(output, yResolutionOffset, options.DpiY);
        }
    }
}
