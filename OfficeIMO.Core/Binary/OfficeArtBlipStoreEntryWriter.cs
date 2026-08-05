using System;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Writes bounded OfficeArt File BLIP Store Entry and image BLIP records.</summary>
public static class OfficeArtBlipStoreEntryWriter {
    private const ushort OfficeArtFbse = 0xF007;
    private const int MaximumImageBytes = 64 * 1024 * 1024;

    /// <summary>
    /// Creates one complete embedded FBSE record for a supported raster or
    /// Windows metafile image.
    /// </summary>
    public static byte[] CreateEmbedded(byte[] imageBytes, string contentType,
        uint referenceCount = 1) {
        ValidateReferenceCount(referenceCount);
        PreparedBlip prepared = PrepareBlip(imageBytes, contentType);
        var fbsePayload = BuildFbsePayload(prepared, referenceCount,
            delayedStreamOffset: 0U, embedBlip: true);
        // An embedded BLIP makes the delay offset irrelevant. Zero avoids
        // declaring the special no-delay value that requires an empty slot.
        return BuildRecord(version: 2, prepared.Format.BlipType, OfficeArtFbse,
            fbsePayload);
    }

    /// <summary>
    /// Creates one FBSE record that points to an image BLIP in an associated
    /// <c>OfficeArtBStoreDelay</c> stream.
    /// </summary>
    public static byte[] CreateDelayed(byte[] imageBytes, string contentType,
        uint delayedStreamOffset, uint referenceCount = 1) {
        ValidateReferenceCount(referenceCount);
        if (delayedStreamOffset == uint.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(delayedStreamOffset),
                "A populated delayed BLIP entry requires a concrete stream offset.");
        }
        PreparedBlip prepared = PrepareBlip(imageBytes, contentType);
        byte[] fbsePayload = BuildFbsePayload(prepared, referenceCount,
            delayedStreamOffset, embedBlip: false);
        return BuildRecord(version: 2, prepared.Format.BlipType, OfficeArtFbse,
            fbsePayload);
    }

    /// <summary>
    /// Creates one standalone supported raster or Windows metafile OfficeArt
    /// BLIP record.
    /// </summary>
    public static byte[] CreateBlipRecord(byte[] imageBytes,
        string contentType) => PrepareBlip(imageBytes, contentType).Record;

    private static PreparedBlip PrepareBlip(byte[] imageBytes,
        string contentType) {
        if (imageBytes == null) throw new ArgumentNullException(nameof(imageBytes));
        if (imageBytes.Length == 0) {
            throw new ArgumentException("Image data cannot be empty.",
                nameof(imageBytes));
        }
        if (imageBytes.Length > MaximumImageBytes) {
            throw new ArgumentException(
                $"OfficeArt BLIP payloads cannot exceed {MaximumImageBytes} bytes.",
                nameof(imageBytes));
        }
        if (string.IsNullOrWhiteSpace(contentType)) {
            throw new ArgumentException("Image content type is required.",
                nameof(contentType));
        }

        BlipFormat format = ResolveFormat(contentType, imageBytes);
        byte[] fileData = format.StripBitmapHeader
            ? CopyRange(imageBytes, 14, imageBytes.Length - 14)
            : (byte[])imageBytes.Clone();
        byte[] uid = OfficeArtMd4.Compute(fileData);
        byte[] blipPayload = format.MetafileType == MetafileType.None
            ? BuildRasterPayload(uid, fileData)
            : BuildMetafilePayload(format.MetafileType, uid, fileData);
        byte[] record = BuildRecord(version: 0, format.RecordInstance,
            format.RecordType, blipPayload);
        return new PreparedBlip(format, uid, record);
    }

    private static byte[] BuildRasterPayload(byte[] uid, byte[] fileData) {
        var payload = new byte[checked(17 + fileData.Length)];
        Buffer.BlockCopy(uid, 0, payload, 0, uid.Length);
        payload[16] = 0xFF;
        Buffer.BlockCopy(fileData, 0, payload, 17, fileData.Length);
        return payload;
    }

    private static byte[] BuildMetafilePayload(MetafileType type,
        byte[] uid, byte[] fileData) {
        MetafileBounds bounds = type == MetafileType.Emf
            ? ReadEmfBounds(fileData)
            : ReadWmfBounds(fileData);
        byte[] storedData = OfficeZlibCodec.Compress(fileData);
        var payload = new byte[checked(50 + storedData.Length)];
        Buffer.BlockCopy(uid, 0, payload, 0, uid.Length);
        WriteUInt32(payload, 16, checked((uint)fileData.Length));
        WriteUInt32(payload, 20, unchecked((uint)bounds.Left));
        WriteUInt32(payload, 24, unchecked((uint)bounds.Top));
        WriteUInt32(payload, 28, unchecked((uint)bounds.Right));
        WriteUInt32(payload, 32, unchecked((uint)bounds.Bottom));
        WriteUInt32(payload, 36, unchecked((uint)bounds.WidthEmus));
        WriteUInt32(payload, 40, unchecked((uint)bounds.HeightEmus));
        WriteUInt32(payload, 44, checked((uint)storedData.Length));
        payload[48] = 0x00;
        payload[49] = 0xFE;
        Buffer.BlockCopy(storedData, 0, payload, 50, storedData.Length);
        return payload;
    }

    private static MetafileBounds ReadEmfBounds(byte[] data) {
        if (data.Length < 88 || ReadUInt32(data, 0) != 1U
            || ReadUInt32(data, 4) < 88U
            || ReadUInt32(data, 40) != 0x464D4520U) {
            throw new NotSupportedException(
                "The EMF payload has no valid enhanced-metafile header.");
        }
        int left = ReadInt32(data, 8);
        int top = ReadInt32(data, 12);
        int right = ReadInt32(data, 16);
        int bottom = ReadInt32(data, 20);
        int frameLeft = ReadInt32(data, 24);
        int frameTop = ReadInt32(data, 28);
        int frameRight = ReadInt32(data, 32);
        int frameBottom = ReadInt32(data, 36);
        long widthEmus = checked(((long)frameRight - frameLeft) * 360L);
        long heightEmus = checked(((long)frameBottom - frameTop) * 360L);
        return CreateMetafileBounds(left, top, right, bottom, widthEmus,
            heightEmus, "EMF");
    }

    private static MetafileBounds ReadWmfBounds(byte[] data) {
        if (data.Length < 40 || ReadUInt32(data, 0) != 0x9AC6CDD7U) {
            throw new NotSupportedException(
                "The WMF payload requires a valid placeable-metafile header.");
        }
        int left = unchecked((short)ReadUInt16(data, 6));
        int top = unchecked((short)ReadUInt16(data, 8));
        int right = unchecked((short)ReadUInt16(data, 10));
        int bottom = unchecked((short)ReadUInt16(data, 12));
        ushort unitsPerInch = ReadUInt16(data, 14);
        if (unitsPerInch == 0) {
            throw new NotSupportedException(
                "The placeable WMF payload declares zero units per inch.");
        }
        long widthEmus = checked((long)Math.Round(
            (right - left) * 914400D / unitsPerInch,
            MidpointRounding.AwayFromZero));
        long heightEmus = checked((long)Math.Round(
            (bottom - top) * 914400D / unitsPerInch,
            MidpointRounding.AwayFromZero));
        return CreateMetafileBounds(left, top, right, bottom, widthEmus,
            heightEmus, "WMF");
    }

    private static MetafileBounds CreateMetafileBounds(int left, int top,
        int right, int bottom, long widthEmus, long heightEmus,
        string formatName) {
        if (right <= left || bottom <= top || widthEmus <= 0
            || heightEmus <= 0 || widthEmus > int.MaxValue
            || heightEmus > int.MaxValue) {
            throw new NotSupportedException(
                $"The {formatName} payload has invalid or oversized bounds.");
        }
        return new MetafileBounds(left, top, right, bottom,
            checked((int)widthEmus), checked((int)heightEmus));
    }

    private static byte[] BuildFbsePayload(PreparedBlip prepared,
        uint referenceCount, uint delayedStreamOffset, bool embedBlip) {
        int embeddedLength = embedBlip ? prepared.Record.Length : 0;
        var payload = new byte[checked(36 + embeddedLength)];
        payload[0] = prepared.Format.BlipType;
        payload[1] = prepared.Format.MacOsBlipType;
        Buffer.BlockCopy(prepared.Uid, 0, payload, 2, prepared.Uid.Length);
        WriteUInt16(payload, 18, 0x00FF);
        WriteUInt32(payload, 20, checked((uint)prepared.Record.Length));
        WriteUInt32(payload, 24, referenceCount);
        WriteUInt32(payload, 28, delayedStreamOffset);
        if (embedBlip) {
            Buffer.BlockCopy(prepared.Record, 0, payload, 36,
                prepared.Record.Length);
        }
        return payload;
    }

    private static void ValidateReferenceCount(uint referenceCount) {
        if (referenceCount == 0) {
            throw new ArgumentOutOfRangeException(nameof(referenceCount));
        }
    }

    private static BlipFormat ResolveFormat(string contentType,
        byte[] imageBytes) {
        string normalized = contentType.Trim().ToLowerInvariant();
        if (normalized is "image/png" or "image/x-png"
            && HasPrefix(imageBytes, 0x89, 0x50, 0x4E, 0x47)) {
            return new BlipFormat(0x06, 0x06E0, 0xF01E,
                stripBitmapHeader: false);
        }
        if (normalized is "image/jpeg" or "image/jpg"
            && HasPrefix(imageBytes, 0xFF, 0xD8)) {
            return new BlipFormat(0x05, 0x046A, 0xF01D,
                stripBitmapHeader: false);
        }
        if (normalized is "image/bmp" or "image/x-ms-bmp"
            && imageBytes.Length > 26
            && imageBytes[0] == (byte)'B' && imageBytes[1] == (byte)'M') {
            return new BlipFormat(0x07, 0x07A8, 0xF01F,
                stripBitmapHeader: true);
        }
        if (normalized is "image/tiff" or "image/tif"
            && (HasPrefix(imageBytes, 0x49, 0x49, 0x2A, 0x00)
                || HasPrefix(imageBytes, 0x4D, 0x4D, 0x00, 0x2A))) {
            return new BlipFormat(0x11, 0x06E4, 0xF029,
                stripBitmapHeader: false);
        }
        if (normalized is "image/x-emf" or "image/emf"
            && imageBytes.Length >= 44
            && ReadUInt32(imageBytes, 0) == 1U
            && ReadUInt32(imageBytes, 40) == 0x464D4520U) {
            return new BlipFormat(0x02, 0x03D4, 0xF01A,
                stripBitmapHeader: false, MetafileType.Emf,
                macOsBlipType: 0x04);
        }
        if (normalized is "image/x-wmf" or "image/wmf"
            && imageBytes.Length >= 4
            && ReadUInt32(imageBytes, 0) == 0x9AC6CDD7U) {
            return new BlipFormat(0x03, 0x0216, 0xF01B,
                stripBitmapHeader: false, MetafileType.Wmf,
                macOsBlipType: 0x04);
        }
        throw new NotSupportedException(
            $"OfficeArt BLIP writing does not support '{contentType}' or its payload signature is invalid.");
    }

    private static bool HasPrefix(byte[] source, params byte[] prefix) {
        if (source.Length < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (source[index] != prefix[index]) return false;
        }
        return true;
    }

    private static byte[] CopyRange(byte[] source, int offset, int count) {
        var result = new byte[count];
        Buffer.BlockCopy(source, offset, result, 0, count);
        return result;
    }

    private static byte[] BuildRecord(byte version, ushort instance,
        ushort type, byte[] payload) {
        var result = new byte[checked(8 + payload.Length)];
        WriteUInt16(result, 0,
            unchecked((ushort)(instance << 4 | version)));
        WriteUInt16(result, 2, type);
        WriteUInt32(result, 4, checked((uint)payload.Length));
        Buffer.BlockCopy(payload, 0, result, 8, payload.Length);
        return result;
    }

    private static void WriteUInt16(byte[] target, int offset,
        ushort value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
    }

    private static void WriteUInt32(byte[] target, int offset, uint value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
        target[offset + 2] = unchecked((byte)(value >> 16));
        target[offset + 3] = unchecked((byte)(value >> 24));
    }

    private static ushort ReadUInt16(byte[] source, int offset) =>
        unchecked((ushort)(source[offset] | source[offset + 1] << 8));

    private static uint ReadUInt32(byte[] source, int offset) =>
        unchecked((uint)(source[offset]
            | source[offset + 1] << 8
            | source[offset + 2] << 16
            | source[offset + 3] << 24));

    private static int ReadInt32(byte[] source, int offset) =>
        unchecked((int)ReadUInt32(source, offset));

    private readonly struct BlipFormat {
        internal BlipFormat(byte blipType, ushort recordInstance,
            ushort recordType, bool stripBitmapHeader,
            MetafileType metafileType = MetafileType.None,
            byte? macOsBlipType = null) {
            BlipType = blipType;
            RecordInstance = recordInstance;
            RecordType = recordType;
            StripBitmapHeader = stripBitmapHeader;
            MetafileType = metafileType;
            MacOsBlipType = macOsBlipType ?? blipType;
        }

        internal byte BlipType { get; }
        internal ushort RecordInstance { get; }
        internal ushort RecordType { get; }
        internal bool StripBitmapHeader { get; }
        internal MetafileType MetafileType { get; }
        internal byte MacOsBlipType { get; }
    }

    private readonly struct PreparedBlip {
        internal PreparedBlip(BlipFormat format, byte[] uid,
            byte[] record) {
            Format = format;
            Uid = uid;
            Record = record;
        }

        internal BlipFormat Format { get; }
        internal byte[] Uid { get; }
        internal byte[] Record { get; }
    }

    private enum MetafileType {
        None,
        Emf,
        Wmf
    }

    private readonly struct MetafileBounds {
        internal MetafileBounds(int left, int top, int right, int bottom,
            int widthEmus, int heightEmus) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            WidthEmus = widthEmus;
            HeightEmus = heightEmus;
        }

        internal int Left { get; }
        internal int Top { get; }
        internal int Right { get; }
        internal int Bottom { get; }
        internal int WidthEmus { get; }
        internal int HeightEmus { get; }
    }
}
