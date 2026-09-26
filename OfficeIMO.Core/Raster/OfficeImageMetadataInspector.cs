using System;
using System.Threading;
namespace OfficeIMO.Drawing;

internal sealed class OfficeImageMetadataSnapshot {
    internal OfficeImageMetadataKinds Kinds { get; set; }
    internal bool HasColorRenderingMetadata { get; set; }
    internal byte[]? Exif { get; set; }
    internal byte[]? Xmp { get; set; }
    internal byte[]? Icc { get; set; }
    internal bool HasDuplicateJpegExif { get; set; }
    internal bool HasExtendedJpegXmp { get; set; }
    internal bool HasDuplicateStandardJpegXmp { get; set; }
    internal bool ExifContainsResolution { get; set; }
    internal bool HasPhysicalResolution { get; set; }
    internal bool HasUnitlessResolution { get; set; }
    internal double? PhysicalDpiX { get; set; }
    internal double? PhysicalDpiY { get; set; }
}

internal static partial class OfficeImageMetadataInspector {
    private static readonly byte[] ExifPrefix = { (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0 };
    private static readonly byte[] XmpPrefix = System.Text.Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
    private static readonly byte[] ExtendedXmpPrefix = System.Text.Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
    private static readonly byte[] IccPrefix = System.Text.Encoding.ASCII.GetBytes("ICC_PROFILE\0");

    internal static OfficeImageMetadataSnapshot Inspect(
        byte[] data,
        OfficeImageFormat format,
        long retainedManagedBytes = 0L) =>
        Inspect(data, format, retainedManagedBytes, CancellationToken.None);

    internal static OfficeImageMetadataSnapshot Inspect(
        byte[] data,
        OfficeImageFormat format,
        long retainedManagedBytes,
        CancellationToken cancellationToken) {
        if (retainedManagedBytes < 0L) throw new ArgumentOutOfRangeException(nameof(retainedManagedBytes));
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = new OfficeImageMetadataSnapshot();
        if (OfficeImageOrientationNormalizer.TryRead(data, cancellationToken, out OfficeImageOrientation orientation) &&
            orientation != OfficeImageOrientation.Normal) snapshot.Kinds |= OfficeImageMetadataKinds.Orientation;
        switch (format) {
            case OfficeImageFormat.Jpeg:
                InspectJpeg(data, snapshot, retainedManagedBytes, cancellationToken);
                break;
            case OfficeImageFormat.Png:
                InspectPng(data, snapshot, cancellationToken);
                break;
            case OfficeImageFormat.Webp:
                InspectWebp(data, snapshot, cancellationToken);
                break;
            case OfficeImageFormat.Tiff:
                InspectTiff(data, snapshot, cancellationToken);
                break;
            case OfficeImageFormat.Bmp:
                InspectBmp(data, snapshot, cancellationToken);
                break;
            case OfficeImageFormat.Gif:
                InspectGif(data, snapshot, cancellationToken);
                break;
        }
        return snapshot;
    }

    internal static void InspectJpeg(
        byte[] data,
        OfficeImageMetadataSnapshot snapshot,
        long retainedManagedBytes,
        CancellationToken cancellationToken) {
        JpegIccPart?[]? iccParts = null;
        bool invalidIccSequence = false;
        int offset = 2;
        bool inScan = false;
        long nextCancellationOffset = 0L;
        while (offset < data.Length - 1) {
            if (offset >= nextCancellationOffset) {
                cancellationToken.ThrowIfCancellationRequested();
                nextCancellationOffset = (long)offset + 64L * 1024L;
            }
            if (inScan && data[offset] != 0xFF) {
                offset++;
                continue;
            }
            if (data[offset++] != 0xFF) break;
            while (offset < data.Length && data[offset] == 0xFF) {
                offset++;
                if (offset >= nextCancellationOffset) {
                    cancellationToken.ThrowIfCancellationRequested();
                    nextCancellationOffset = (long)offset + 64L * 1024L;
                }
            }
            if (offset >= data.Length) break;
            int marker = data[offset++];
            if (inScan) {
                if (marker == 0x00 || marker >= 0xD0 && marker <= 0xD7) continue;
                inScan = false;
            }
            if (marker == 0xD9) break;
            if (marker >= 0xD0 && marker <= 0xD7 || marker == 0x01) continue;
            if (offset > data.Length - 2) break;
            int length = data[offset] << 8 | data[offset + 1];
            if (length < 2 || offset > data.Length - length) break;
            int payload = offset + 2;
            int count = length - 2;
            if (marker == 0xE0 && Matches(data, payload, count, "JFIF\0")) {
                bool physical = count >= 12 && data[payload + 7] >= 1 && data[payload + 7] <= 2;
                MarkResolution(snapshot, physical);
                if (physical) {
                    int densityX = data[payload + 8] << 8 | data[payload + 9];
                    int densityY = data[payload + 10] << 8 | data[payload + 11];
                    double scale = data[payload + 7] == 2 ? 2.54D : 1D;
                    SetPhysicalResolution(snapshot, densityX * scale, densityY * scale, overwrite: true);
                }
            }
            if (marker == 0xE1 && StartsWith(data, payload, count, ExifPrefix)) {
                if (snapshot.Exif != null) snapshot.HasDuplicateJpegExif = true;
                snapshot.Exif = Slice(data, payload, count, cancellationToken);
                InspectExifPayload(snapshot.Exif, 0, snapshot.Exif.Length, snapshot, cancellationToken);
            } else if (marker == 0xE1 && StartsWith(data, payload, count, XmpPrefix)) {
                if (snapshot.Xmp != null) snapshot.HasDuplicateStandardJpegXmp = true;
                snapshot.Xmp = Slice(data, payload, count, cancellationToken);
                snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            } else if (marker == 0xE1 && StartsWith(data, payload, count, ExtendedXmpPrefix)) {
                snapshot.HasExtendedJpegXmp = true;
                snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            } else if (marker == 0xE2 && StartsWith(data, payload, count, IccPrefix) && count >= IccPrefix.Length + 2) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                snapshot.HasColorRenderingMetadata = true;
                int sequence = data[payload + IccPrefix.Length];
                int total = data[payload + IccPrefix.Length + 1];
                if (sequence < 1 || total < 1 || sequence > total) {
                    invalidIccSequence = true;
                } else {
                    if (iccParts == null) iccParts = new JpegIccPart?[total];
                    if (iccParts.Length != total || iccParts[sequence - 1].HasValue) {
                        invalidIccSequence = true;
                    } else {
                        iccParts[sequence - 1] = new JpegIccPart(
                            payload + IccPrefix.Length + 2,
                            count - IccPrefix.Length - 2);
                    }
                }
            } else if (marker == 0xE2 && StartsWith(data, payload, count, IccPrefix)) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                snapshot.HasColorRenderingMetadata = true;
                invalidIccSequence = true;
            } else if (marker == 0xFE) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            }
            if (marker == 0xDA) inScan = true;
            offset += length;
        }
        if (!invalidIccSequence && iccParts != null) {
            int length = 0;
            for (int index = 0; index < iccParts.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!iccParts[index].HasValue) return;
                JpegIccPart part = iccParts[index]!.Value;
                if (part.Length > OfficeRasterGuards.MaximumEncodedBytes - length) return;
                length += part.Length;
            }
            long inspectorRetainedBytes;
            try {
                inspectorRetainedBytes = checked(
                    retainedManagedBytes +
                    (snapshot.Exif == null ? 0L : snapshot.Exif.LongLength + 24L) +
                    (snapshot.Xmp == null ? 0L : snapshot.Xmp.LongLength + 24L) +
                    24L + iccParts.LongLength * 24L);
            } catch (OverflowException) {
                return;
            }
            if (!IsJpegIccAssemblyWithinLimit(data.LongLength, length, inspectorRetainedBytes)) return;
            snapshot.Icc = new byte[length];
            int target = 0;
            for (int index = 0; index < iccParts.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                JpegIccPart part = iccParts[index]!.Value;
                CopyWithCancellation(data, part.Offset, snapshot.Icc, target, part.Length, cancellationToken);
                target += part.Length;
            }
        }
    }

    private static void InspectPng(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken cancellationToken) {
        int offset = 8;
        bool hasStandardRgb = false;
        bool hasGamma = false;
        bool hasStandardGamma = false;
        bool hasChromaticities = false;
        bool hasStandardChromaticities = false;
        while (offset <= data.Length - 12) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadBigEndian(data, offset);
            if (length < 0 || offset > data.Length - 12 - length) break;
            string type = ReadAscii(data, offset + 4, 4);
            if (type == "eXIf") {
                InspectExifPayload(data, offset + 8, length, snapshot, cancellationToken);
            } else if (type == "iCCP") {
                snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                snapshot.HasColorRenderingMetadata = true;
            } else if (type == "gAMA") {
                hasGamma = true;
                hasStandardGamma = length == 4 && ReadUInt32Unsigned(data, offset + 8, little: false) == 45455U;
            } else if (type == "cHRM") {
                hasChromaticities = true;
                hasStandardChromaticities = length == 32 &&
                    OfficePngContainerValidator.HasStandardRgbChromaticities(data, offset + 8);
            } else if (type == "sRGB") hasStandardRgb = true;
            else if (type == "cICP") {
                bool canonicalSrgb = length == 4 &&
                    data[offset + 8] == 1 &&
                    data[offset + 9] == 13 &&
                    data[offset + 10] == 0 &&
                    data[offset + 11] == 1;
                if (!canonicalSrgb) snapshot.HasColorRenderingMetadata = true;
            }
            else if (type == "pHYs") {
                bool physical = length == 9 && data[offset + 16] == 1;
                MarkResolution(snapshot, physical);
                if (physical) {
                    const double pixelsPerMeterPerDpi = 39.37007874015748D;
                    SetPhysicalResolution(snapshot,
                        ReadUInt32Unsigned(data, offset + 8, little: false) / pixelsPerMeterPerDpi,
                        ReadUInt32Unsigned(data, offset + 12, little: false) / pixelsPerMeterPerDpi,
                        overwrite: true);
                }
            } else if ((type == "tEXt" || type == "zTXt" || type == "iTXt") &&
                       HasExactPngTextKeyword(
                           data, offset + 8, length, "XML:com.adobe.xmp")) snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (type == "tEXt" || type == "zTXt" || type == "iTXt") snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            offset = checked(offset + 12 + length);
        }
        // Matching gAMA/cHRM values alone do not declare the exact sRGB transfer
        // function. The decoder leaves those chunks unapplied, so classification
        // can only use the original bytes when an sRGB declaration is present.
        if ((hasGamma || hasChromaticities) &&
            (!hasStandardRgb || hasGamma && !hasStandardGamma || hasChromaticities && !hasStandardChromaticities)) {
            snapshot.HasColorRenderingMetadata = true;
        }
    }

    private static void InspectWebp(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken cancellationToken) {
        int offset = 12;
        while (offset <= data.Length - 8) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadLittleEndian(data, offset + 4);
            if (length < 0 || offset > data.Length - 8 - length) break;
            string type = ReadAscii(data, offset, 4);
            if (type == "EXIF") {
                InspectExifPayload(data, offset + 8, length, snapshot, cancellationToken);
            } else if (type == "XMP ") snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (type == "ICCP") {
                snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                snapshot.HasColorRenderingMetadata = true;
            }
            offset = checked(offset + 8 + length + (length & 1));
        }
    }

    private static void InspectTiff(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length < 10) return;
        bool little = data[0] == (byte)'I';
        int ifd = ReadUInt32(data, 4, little);
        if (ifd < 0 || ifd > data.Length - 2) return;
        int count = ReadUInt16(data, ifd, little);
        bool hasResolution = false;
        int resolutionUnit = 2;
        double? resolutionX = null;
        double? resolutionY = null;
        int bitsPerSampleEntry = -1;
        int photometricInterpretation = -1;
        int samplesPerPixel = -1;
        int transferFunctionEntry = -1;
        int whitePointEntry = -1;
        int primaryChromaticitiesEntry = -1;
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int entry = ifd + 2 + index * 12;
            if (entry > data.Length - 12) return;
            int tag = ReadUInt16(data, entry, little);
            if (tag == 34665) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Exif;
                InspectExifSubIfdColorMetadata(
                    data, entry, little, 0, data.Length, snapshot, cancellationToken);
            } else if (tag == 700) snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (tag == 34675) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                snapshot.HasColorRenderingMetadata = true;
            } else if (tag == 270) snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            else if (tag == 258) bitsPerSampleEntry = entry;
            else if (tag == 262) {
                if (!TryReadInlineShort(data, entry, little, data.Length, out photometricInterpretation)) {
                    snapshot.HasColorRenderingMetadata = true;
                }
            } else if (tag == 277) {
                if (!TryReadInlineShort(data, entry, little, data.Length, out samplesPerPixel)) {
                    snapshot.HasColorRenderingMetadata = true;
                }
            } else if (tag == 301) {
                if (transferFunctionEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                transferFunctionEntry = entry;
            } else if (tag == 318) {
                if (whitePointEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                whitePointEntry = entry;
            } else if (tag == 319) {
                if (primaryChromaticitiesEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                primaryChromaticitiesEntry = entry;
            } else if (IsUnappliedTiffColorTag(tag)) snapshot.HasColorRenderingMetadata = true;
            else if (tag == 282 || tag == 283) {
                hasResolution = true;
                if (TryReadRational(
                        data, entry, little, tiffBaseOffset: 0, data.Length, out double resolution)) {
                    if (tag == 282) resolutionX = resolution;
                    else resolutionY = resolution;
                }
            } else if (tag == 296) {
                hasResolution = true;
                if (!TryReadInlineShort(data, entry, little, data.Length, out resolutionUnit)) resolutionUnit = 1;
            }
        }
        if (hasResolution) {
            bool physical = resolutionUnit == 2 || resolutionUnit == 3;
            MarkResolution(snapshot, physical);
            if (physical && resolutionX.HasValue && resolutionY.HasValue) {
                double scale = resolutionUnit == 3 ? 2.54D : 1D;
                SetPhysicalResolution(snapshot, resolutionX.Value * scale, resolutionY.Value * scale, overwrite: true);
            }
        }
        bool hasTiffColorimetry = transferFunctionEntry >= 0 || whitePointEntry >= 0 || primaryChromaticitiesEntry >= 0;
        if (hasTiffColorimetry && !IsCanonicalSrgbTiffColorimetry(
                data,
                little,
                tiffBaseOffset: 0,
                data.Length,
                requireRgbImageTags: true,
                hasSrgbColorSpace: false,
                bitsPerSampleEntry,
                photometricInterpretation,
                samplesPerPixel,
                transferFunctionEntry,
                whitePointEntry,
                primaryChromaticitiesEntry,
                cancellationToken)) {
            snapshot.HasColorRenderingMetadata = true;
        }
    }

    internal static bool IsJpegIccAssemblyWithinLimit(
        long encodedBytes,
        long iccBytes,
        long retainedManagedBytes) {
        if (encodedBytes < 0L || iccBytes < 0L || retainedManagedBytes < 0L) return false;
        try {
            long peakBytes = checked(
                encodedBytes + 24L + iccBytes + 24L + retainedManagedBytes);
            return peakBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool HasExactPngTextKeyword(
        byte[] data,
        int payloadOffset,
        int payloadLength,
        string keyword) {
        if (payloadLength <= keyword.Length || payloadOffset < 0 ||
            payloadOffset > data.Length - payloadLength) return false;
        for (int index = 0; index < keyword.Length; index++) {
            if (data[payloadOffset + index] != (byte)keyword[index]) return false;
        }
        return data[payloadOffset + keyword.Length] == 0;
    }

    private static void InspectBmp(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        const int dibHeaderOffset = 14;
        const int minimumInfoHeaderSize = 40;
        const int horizontalPixelsPerMeterOffset = 38;
        const int verticalPixelsPerMeterOffset = 42;
        if (data.Length < verticalPixelsPerMeterOffset + 4 ||
            ReadLittleEndian(data, dibHeaderOffset) < minimumInfoHeaderSize) return;

        int dibHeaderSize = ReadLittleEndian(data, dibHeaderOffset);
        if (dibHeaderSize >= 108 && data.Length >= dibHeaderOffset + 60) {
            const int logicalColorSpaceSrgb = 0x73524742;
            int colorSpaceType = ReadLittleEndian(data, dibHeaderOffset + 56);
            if (colorSpaceType != logicalColorSpaceSrgb) snapshot.HasColorRenderingMetadata = true;
        }

        int horizontalPixelsPerMeter = ReadLittleEndian(data, horizontalPixelsPerMeterOffset);
        int verticalPixelsPerMeter = ReadLittleEndian(data, verticalPixelsPerMeterOffset);
        if (horizontalPixelsPerMeter > 0 || verticalPixelsPerMeter > 0) {
            MarkResolution(snapshot, isPhysical: true);
            if (horizontalPixelsPerMeter > 0 && verticalPixelsPerMeter > 0) {
                const double pixelsPerMeterPerDpi = 39.37007874015748D;
                SetPhysicalResolution(snapshot,
                    horizontalPixelsPerMeter / pixelsPerMeterPerDpi,
                    verticalPixelsPerMeter / pixelsPerMeterPerDpi,
                    overwrite: true);
            }
        }
    }

    private static void InspectExifPayload(
        byte[] exif,
        int payloadOffset,
        int payloadLength,
        OfficeImageMetadataSnapshot snapshot,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        snapshot.Kinds |= OfficeImageMetadataKinds.Exif;
        if (OfficeImageOrientationNormalizer.TryReadExifOrientationPayload(
                exif, payloadOffset, payloadLength, out OfficeImageOrientation orientation) &&
            orientation != OfficeImageOrientation.Normal) {
            snapshot.Kinds |= OfficeImageMetadataKinds.Orientation;
        }
        if (payloadOffset < 0 || payloadLength < 0 || payloadOffset > exif.Length - payloadLength) return;
        int payloadEnd = payloadOffset + payloadLength;
        int tiffOffset = payloadLength >= 6 && StartsWith(exif, payloadOffset, payloadLength, ExifPrefix)
            ? payloadOffset + 6
            : payloadOffset;
        if (payloadEnd - tiffOffset < 10) return;
        bool little = exif[tiffOffset] == (byte)'I' && exif[tiffOffset + 1] == (byte)'I';
        bool big = exif[tiffOffset] == (byte)'M' && exif[tiffOffset + 1] == (byte)'M';
        if (!little && !big) return;
        int ifd = ReadUInt32(exif, tiffOffset + 4, little);
        if (ifd < 0 || ifd > payloadEnd - tiffOffset - 2) return;
        int absoluteIfd = tiffOffset + ifd;
        int count = ReadUInt16(exif, absoluteIfd, little);
        bool hasResolution = false;
        int resolutionUnit = 2;
        double? resolutionX = null;
        double? resolutionY = null;
        int transferFunctionEntry = -1;
        int whitePointEntry = -1;
        int primaryChromaticitiesEntry = -1;
        bool hasSrgbColorSpace = false;
        for (int index = 0; index < count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            int entry = absoluteIfd + 2 + index * 12;
            if (entry > payloadEnd - 12) {
                if (transferFunctionEntry >= 0 || whitePointEntry >= 0 || primaryChromaticitiesEntry >= 0) {
                    snapshot.HasColorRenderingMetadata = true;
                }
                return;
            }
            int tag = ReadUInt16(exif, entry, little);
            if (tag == 301) {
                if (transferFunctionEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                transferFunctionEntry = entry;
            } else if (tag == 318) {
                if (whitePointEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                whitePointEntry = entry;
            } else if (tag == 319) {
                if (primaryChromaticitiesEntry >= 0) snapshot.HasColorRenderingMetadata = true;
                primaryChromaticitiesEntry = entry;
            } else if (tag == 34665) {
                hasSrgbColorSpace |= InspectExifSubIfdColorMetadata(
                    exif, entry, little, tiffOffset, payloadEnd, snapshot, cancellationToken);
            } else if (IsUnappliedTiffColorTag(tag)) {
                snapshot.HasColorRenderingMetadata = true;
            } else if (tag == 282 || tag == 283) {
                hasResolution = true;
                if (TryReadRational(
                        exif, entry, little, tiffOffset, payloadEnd, out double resolution)) {
                    if (tag == 282) resolutionX = resolution;
                    else resolutionY = resolution;
                }
            } else if (tag == 296) {
                hasResolution = true;
                if (!TryReadInlineShort(
                        exif, entry, little, payloadEnd, out resolutionUnit)) resolutionUnit = 1;
            }
        }
        if (hasResolution) {
            snapshot.ExifContainsResolution = true;
            bool physical = resolutionUnit == 2 || resolutionUnit == 3;
            MarkResolution(snapshot, physical);
            if (physical && resolutionX.HasValue && resolutionY.HasValue) {
                double scale = resolutionUnit == 3 ? 2.54D : 1D;
                SetPhysicalResolution(snapshot, resolutionX.Value * scale, resolutionY.Value * scale, overwrite: false);
            }
        }
        bool hasTiffColorimetry = transferFunctionEntry >= 0 || whitePointEntry >= 0 || primaryChromaticitiesEntry >= 0;
        if (hasTiffColorimetry && !IsCanonicalSrgbTiffColorimetry(
                exif,
                little,
                tiffOffset,
                payloadEnd,
                requireRgbImageTags: false,
                hasSrgbColorSpace,
                bitsPerSampleEntry: -1,
                photometricInterpretation: -1,
                samplesPerPixel: -1,
                transferFunctionEntry,
                whitePointEntry,
                primaryChromaticitiesEntry,
                cancellationToken)) {
            snapshot.HasColorRenderingMetadata = true;
        }
    }

    private static bool InspectExifSubIfdColorMetadata(
        byte[] data,
        int pointerEntry,
        bool little,
        int tiffBaseOffset,
        int viewEnd,
        OfficeImageMetadataSnapshot snapshot,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pointerEntry < 0 || pointerEntry > viewEnd - 12 ||
            ReadUInt16(data, pointerEntry + 2, little) != 4 ||
            ReadUInt32Unsigned(data, pointerEntry + 4, little) != 1U) return false;
        uint relativeOffset = ReadUInt32Unsigned(data, pointerEntry + 8, little);
        long absoluteOffset = (long)tiffBaseOffset + relativeOffset;
        if (absoluteOffset < tiffBaseOffset || absoluteOffset > viewEnd - 2) return false;
        int subIfd = (int)absoluteOffset;
        int count = ReadUInt16(data, subIfd, little);
        bool hasSrgbColorSpace = false;
        for (int index = 0; index < count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            int entry = subIfd + 2 + index * 12;
            if (entry < tiffBaseOffset || entry > viewEnd - 12) {
                if (hasSrgbColorSpace) snapshot.HasColorRenderingMetadata = true;
                return false;
            }
            int tag = ReadUInt16(data, entry, little);
            if (IsUnappliedTiffColorTag(tag)) {
                snapshot.HasColorRenderingMetadata = true;
            } else if (tag == 40961) {
                if (TryReadInlineShort(data, entry, little, viewEnd, out int colorSpace) && colorSpace == 1) {
                    hasSrgbColorSpace = true;
                } else {
                    snapshot.HasColorRenderingMetadata = true;
                }
            }
        }
        return hasSrgbColorSpace;
    }

    private static bool IsUnappliedTiffColorTag(int tag) =>
        tag == 301 || tag == 318 || tag == 319 || tag == 529 || tag == 532 || tag == 34675 || tag == 42240;

    private static void MarkResolution(OfficeImageMetadataSnapshot snapshot, bool isPhysical) {
        snapshot.Kinds |= OfficeImageMetadataKinds.Resolution;
        if (isPhysical) snapshot.HasPhysicalResolution = true;
        else snapshot.HasUnitlessResolution = true;
    }

    private static void SetPhysicalResolution(
        OfficeImageMetadataSnapshot snapshot,
        double dpiX,
        double dpiY,
        bool overwrite) {
        if (double.IsNaN(dpiX) || double.IsInfinity(dpiX) ||
            double.IsNaN(dpiY) || double.IsInfinity(dpiY) || dpiX <= 0D || dpiY <= 0D) return;
        if (overwrite || !snapshot.PhysicalDpiX.HasValue) snapshot.PhysicalDpiX = dpiX;
        if (overwrite || !snapshot.PhysicalDpiY.HasValue) snapshot.PhysicalDpiY = dpiY;
    }

    private static bool TryReadRational(
        byte[] data,
        int entry,
        bool little,
        int tiffBaseOffset,
        int viewEnd,
        out double value) {
        value = 0D;
        if (viewEnd < 0 || viewEnd > data.Length || entry < 0 || entry > viewEnd - 12 ||
            ReadUInt16(data, entry + 2, little) != 5 ||
            ReadUInt32Unsigned(data, entry + 4, little) != 1U) return false;
        uint relativeOffset = ReadUInt32Unsigned(data, entry + 8, little);
        long absoluteOffset = (long)tiffBaseOffset + relativeOffset;
        if (absoluteOffset < 0 || absoluteOffset > viewEnd - 8) return false;
        uint numerator = ReadUInt32Unsigned(data, (int)absoluteOffset, little);
        uint denominator = ReadUInt32Unsigned(data, (int)absoluteOffset + 4, little);
        if (denominator == 0U) return false;
        value = numerator / (double)denominator;
        return !double.IsNaN(value) && !double.IsInfinity(value) && value > 0D;
    }

    private static bool TryReadInlineShort(
        byte[] data,
        int entry,
        bool little,
        int viewEnd,
        out int value) {
        value = 0;
        if (viewEnd < 0 || viewEnd > data.Length || entry < 0 || entry > viewEnd - 12 ||
            ReadUInt16(data, entry + 2, little) != 3 ||
            ReadUInt32(data, entry + 4, little) != 1) return false;
        value = ReadUInt16(data, entry + 8, little);
        return true;
    }

    private static bool StartsWith(byte[] data, int offset, int count, byte[] prefix) {
        if (count < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) if (data[offset + index] != prefix[index]) return false;
        return true;
    }
    private static byte[] Slice(byte[] data, int offset, int count, CancellationToken cancellationToken) {
        var result = new byte[count];
        CopyWithCancellation(data, offset, result, 0, count, cancellationToken);
        return result;
    }
    private static void CopyWithCancellation(
        byte[] source,
        int sourceOffset,
        byte[] destination,
        int destinationOffset,
        int count,
        CancellationToken cancellationToken) {
        const int chunkSize = 64 * 1024;
        for (int copied = 0; copied < count; copied += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            int chunk = Math.Min(chunkSize, count - copied);
            Buffer.BlockCopy(source, sourceOffset + copied, destination, destinationOffset + copied, chunk);
        }
    }
    private static string ReadAscii(byte[] data, int offset, int count) =>
        System.Text.Encoding.ASCII.GetString(data, offset, count);
    private static bool Matches(byte[] data, int offset, int count, string value) =>
        count >= value.Length && ReadAscii(data, offset, value.Length) == value;
    private static void InspectGif(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length < 14) return;
        int offset = 13;
        int packed = data[10];
        if ((packed & 0x80) != 0) offset += 3 << ((packed & 7) + 1);
        while (offset < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int introducer = data[offset++];
            if (introducer == 0x3B) return;
            if (introducer == 0x21) {
                if (offset >= data.Length) return;
                int label = data[offset++];
                if (label == 0xFE) snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
                if (label == 0xFF &&
                    offset < data.Length &&
                    data[offset] == 11 &&
                    offset <= data.Length - 12 &&
                    Matches(data, offset + 1, 11, "ICCRGBG1012")) {
                    snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
                    snapshot.HasColorRenderingMetadata = true;
                }
                if (!SkipGifSubBlocks(data, ref offset, cancellationToken)) return;
                continue;
            }
            if (introducer != 0x2C || offset > data.Length - 9) return;
            int descriptor = data[offset + 8];
            offset += 9;
            if ((descriptor & 0x80) != 0) offset += 3 << ((descriptor & 7) + 1);
            if (offset >= data.Length) return;
            offset++;
            if (!SkipGifSubBlocks(data, ref offset, cancellationToken)) return;
        }
    }
    private static bool SkipGifSubBlocks(
        byte[] data,
        ref int offset,
        CancellationToken cancellationToken) {
        int blockCount = 0;
        while (offset < data.Length) {
            if ((blockCount++ & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            int length = data[offset++];
            if (length == 0) return true;
            if (offset > data.Length - length) return false;
            offset += length;
        }
        return false;
    }
    private static int ReadBigEndian(byte[] data, int offset) =>
        data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3];
    private static int ReadLittleEndian(byte[] data, int offset) =>
        data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24;
    private static int ReadUInt16(byte[] data, int offset, bool little) => little
        ? data[offset] | data[offset + 1] << 8
        : data[offset] << 8 | data[offset + 1];
    private static int ReadUInt32(byte[] data, int offset, bool little) => little
        ? ReadLittleEndian(data, offset)
        : ReadBigEndian(data, offset);
    private static uint ReadUInt32Unsigned(byte[] data, int offset, bool little) => little
        ? (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24)
        : (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);

    private readonly struct JpegIccPart {
        internal JpegIccPart(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }
        internal int Length { get; }
    }
}
