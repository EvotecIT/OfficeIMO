using System;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Writes bounded OfficeArt File BLIP Store Entry and raster BLIP records.</summary>
public static class OfficeArtBlipStoreEntryWriter {
    private const ushort OfficeArtFbse = 0xF007;

    /// <summary>
    /// Creates one complete embedded FBSE record for a PNG, JPEG, BMP/DIB, or TIFF image.
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
    /// Creates one FBSE record that points to a raster BLIP in an associated
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
    /// Creates one standalone PNG, JPEG, BMP/DIB, or TIFF OfficeArt BLIP record.
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
        if (string.IsNullOrWhiteSpace(contentType)) {
            throw new ArgumentException("Image content type is required.",
                nameof(contentType));
        }

        BlipFormat format = ResolveFormat(contentType, imageBytes);
        byte[] fileData = format.StripBitmapHeader
            ? CopyRange(imageBytes, 14, imageBytes.Length - 14)
            : (byte[])imageBytes.Clone();
        byte[] uid = OfficeArtMd4.Compute(fileData);
        var blipPayload = new byte[checked(17 + fileData.Length)];
        Buffer.BlockCopy(uid, 0, blipPayload, 0, uid.Length);
        blipPayload[16] = 0xFF;
        Buffer.BlockCopy(fileData, 0, blipPayload, 17, fileData.Length);
        byte[] record = BuildRecord(version: 0, format.RecordInstance,
            format.RecordType, blipPayload);
        return new PreparedBlip(format, uid, record);
    }

    private static byte[] BuildFbsePayload(PreparedBlip prepared,
        uint referenceCount, uint delayedStreamOffset, bool embedBlip) {
        int embeddedLength = embedBlip ? prepared.Record.Length : 0;
        var payload = new byte[checked(36 + embeddedLength)];
        payload[0] = prepared.Format.BlipType;
        payload[1] = prepared.Format.BlipType;
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
        throw new NotSupportedException(
            $"OfficeArt embedded BLIP writing does not support '{contentType}' or its payload signature is invalid.");
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

    private readonly struct BlipFormat {
        internal BlipFormat(byte blipType, ushort recordInstance,
            ushort recordType, bool stripBitmapHeader) {
            BlipType = blipType;
            RecordInstance = recordInstance;
            RecordType = recordType;
            StripBitmapHeader = stripBitmapHeader;
        }

        internal byte BlipType { get; }
        internal ushort RecordInstance { get; }
        internal ushort RecordType { get; }
        internal bool StripBitmapHeader { get; }
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
}
