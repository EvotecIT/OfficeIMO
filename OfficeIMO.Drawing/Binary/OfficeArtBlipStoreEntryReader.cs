using System;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Reads bounded OfficeArt FBSE records and their embedded or delayed BLIP payloads.</summary>
public static class OfficeArtBlipStoreEntryReader {
    private const int FixedFbseLength = 36;
    private const int DefaultMaximumDecodedImageBytes = 64 * 1024 * 1024;

    /// <summary>
    /// Attempts to decode an FBSE payload. The optional delay stream is used only when no embedded
    /// BLIP follows the FBSE name data. Malformed or oversized image data remains represented as
    /// metadata without returning importable bytes.
    /// </summary>
    public static bool TryRead(byte[] payload, int offset, int length, ushort recordInstance,
        byte[]? delayStream, out OfficeArtBlipStoreEntry? entry,
        int maximumDecodedImageBytes = DefaultMaximumDecodedImageBytes) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        if (offset < 0 || offset > payload.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        if (length < 0 || length > payload.Length - offset) throw new ArgumentOutOfRangeException(nameof(length));
        if (maximumDecodedImageBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumDecodedImageBytes));

        entry = null;
        if (length < FixedFbseLength) return false;
        int endOffset = checked(offset + length);
        byte win32BlipType = payload[offset];
        byte macOsBlipType = payload[offset + 1];
        string uidHex = ToHexString(payload, offset + 2, 16);
        ushort tag = ReadUInt16(payload, offset + 18);
        uint sizeBytes = ReadUInt32(payload, offset + 20);
        uint referenceCount = ReadUInt32(payload, offset + 24);
        uint delayedStreamOffset = ReadUInt32(payload, offset + 28);
        byte nameByteCount = payload[offset + 33];
        string? name = ReadName(payload, offset + FixedFbseLength, endOffset, nameByteCount);

        BlipReadResult blip = default;
        int embeddedOffset = checked(offset + FixedFbseLength + nameByteCount);
        if (embeddedOffset <= endOffset - 8) {
            int embeddedBoundary = GetBlipBoundary(
                embeddedOffset, endOffset, sizeBytes);
            if (embeddedOffset <= embeddedBoundary - 8) {
                blip = ReadBlip(payload, embeddedOffset,
                    embeddedBoundary, OfficeArtBlipStorage.Embedded,
                    maximumDecodedImageBytes);
            }
        }
        if (!blip.HasRecord && delayStream != null && delayedStreamOffset != uint.MaxValue
            && delayedStreamOffset <= int.MaxValue) {
            int delayedOffset = unchecked((int)delayedStreamOffset);
            if (delayedOffset <= delayStream.Length - 8) {
                int delayedBoundary = GetBlipBoundary(
                    delayedOffset, delayStream.Length, sizeBytes);
                if (delayedOffset <= delayedBoundary - 8) {
                    blip = ReadBlip(delayStream, delayedOffset,
                        delayedBoundary, OfficeArtBlipStorage.Delayed,
                        maximumDecodedImageBytes);
                }
            }
        }

        entry = new OfficeArtBlipStoreEntry(recordInstance, win32BlipType, macOsBlipType,
            uidHex, tag, sizeBytes, referenceCount, delayedStreamOffset, nameByteCount, name,
            blip.Storage, blip.RecordVersion, blip.RecordInstance, blip.RecordType,
            blip.PayloadLength, blip.PayloadAvailableLength,
            blip.PayloadSha256, blip.ImageBytes,
            blip.WasImageRejectedBySizeLimit);
        return true;
    }

    /// <summary>Attempts to decode a complete FBSE payload without an associated delay stream.</summary>
    public static bool TryRead(byte[] payload, ushort recordInstance,
        out OfficeArtBlipStoreEntry? entry, int maximumDecodedImageBytes = DefaultMaximumDecodedImageBytes) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        return TryRead(payload, 0, payload.Length, recordInstance, null, out entry,
            maximumDecodedImageBytes);
    }

    internal static bool TryReadBlipRecord(byte[] source, int offset,
        int length, out OfficeArtBlipRecordData data,
        int maximumDecodedImageBytes = DefaultMaximumDecodedImageBytes) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (offset < 0 || offset > source.Length) {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
        if (length < 0 || length > source.Length - offset) {
            throw new ArgumentOutOfRangeException(nameof(length));
        }
        if (maximumDecodedImageBytes < 0) {
            throw new ArgumentOutOfRangeException(
                nameof(maximumDecodedImageBytes));
        }
        data = default;
        if (length < 8) return false;
        BlipReadResult result = ReadBlip(source, offset,
            checked(offset + length), OfficeArtBlipStorage.Embedded,
            maximumDecodedImageBytes);
        if (!result.HasRecord || !result.RecordType.HasValue
            || !result.PayloadLength.HasValue
            || !result.PayloadAvailableLength.HasValue) return false;
        data = new OfficeArtBlipRecordData(result.RecordVersion!.Value,
            result.RecordInstance!.Value, result.RecordType.Value,
            result.PayloadLength.Value, result.PayloadAvailableLength.Value,
            result.PayloadSha256,
            OfficeArtBlipStoreEntry.GetContentType(result.RecordType,
                recordInstanceType: null, win32Type: null, macOsType: null),
            result.ImageBytes ?? Array.Empty<byte>(),
            result.WasImageRejectedBySizeLimit);
        return true;
    }

    internal static bool TryReadBlipStoreFileBlock(byte[] source, int offset,
        int length, out OfficeArtBlipRecordData data,
        int maximumDecodedImageBytes = DefaultMaximumDecodedImageBytes) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (offset < 0 || offset > source.Length) {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
        if (length < 0 || length > source.Length - offset) {
            throw new ArgumentOutOfRangeException(nameof(length));
        }
        data = default;
        if (length < 8) return false;
        ushort versionAndInstance = ReadUInt16(source, offset);
        ushort recordType = ReadUInt16(source, offset + 2);
        if (recordType != 0xF007) {
            return TryReadBlipRecord(source, offset, length, out data,
                maximumDecodedImageBytes);
        }

        uint declaredPayloadLength = ReadUInt32(source, offset + 4);
        int availablePayloadLength = declaredPayloadLength > int.MaxValue
            ? length - 8
            : Math.Min(length - 8, unchecked((int)declaredPayloadLength));
        if (!TryRead(source, offset + 8, availablePayloadLength,
                unchecked((ushort)(versionAndInstance >> 4)), null,
                out OfficeArtBlipStoreEntry? entry,
                maximumDecodedImageBytes)
            || entry == null || !entry.BlipRecordVersion.HasValue
            || !entry.BlipRecordInstance.HasValue
            || !entry.BlipRecordType.HasValue
            || !entry.BlipPayloadLength.HasValue
            || !entry.BlipPayloadAvailableLength.HasValue) return false;
        data = new OfficeArtBlipRecordData(entry.BlipRecordVersion.Value,
            entry.BlipRecordInstance.Value, entry.BlipRecordType.Value,
            entry.BlipPayloadLength.Value,
            entry.BlipPayloadAvailableLength.Value,
            entry.BlipPayloadSha256, entry.ContentType,
            entry.ImageBytes, entry.WasImageRejectedBySizeLimit);
        return true;
    }

    private static int GetBlipBoundary(int recordOffset,
        int sourceBoundary, uint sizeBytes) {
        long declaredBoundary = (long)recordOffset + sizeBytes;
        return declaredBoundary < sourceBoundary
            ? unchecked((int)declaredBoundary)
            : sourceBoundary;
    }

    private static BlipReadResult ReadBlip(byte[] source, int recordOffset, int boundary,
        OfficeArtBlipStorage storage, int maximumDecodedImageBytes) {
        ushort versionAndInstance = ReadUInt16(source, recordOffset);
        byte recordVersion = unchecked((byte)(versionAndInstance & 0x000F));
        ushort recordInstance = unchecked((ushort)(versionAndInstance >> 4));
        ushort recordType = ReadUInt16(source, recordOffset + 2);
        if (!IsBlipRecordType(recordType)) return default;

        uint payloadLength = ReadUInt32(source, recordOffset + 4);
        int payloadOffset = checked(recordOffset + 8);
        int availableLength = payloadLength > int.MaxValue
            ? Math.Max(0, boundary - payloadOffset)
            : Math.Min(Math.Max(0, boundary - payloadOffset), unchecked((int)payloadLength));
        bool isPayloadTruncated =
            payloadLength > unchecked((uint)availableLength);
        string? payloadSha256 = availableLength == 0
            ? null
            : ComputeSha256(source, payloadOffset, availableLength);
        bool wasImageRejectedBySizeLimit = false;
        byte[] imageBytes = availableLength == 0 || isPayloadTruncated
            ? Array.Empty<byte>()
            : ExtractImageBytes(source, payloadOffset, availableLength,
                recordType, recordInstance, maximumDecodedImageBytes,
                out wasImageRejectedBySizeLimit);
        return new BlipReadResult(storage, recordVersion, recordInstance, recordType, payloadLength,
            availableLength, payloadSha256, imageBytes,
            wasImageRejectedBySizeLimit);
    }

    private static byte[] ExtractImageBytes(byte[] source, int offset, int count, ushort recordType,
        ushort recordInstance, int maximumDecodedImageBytes,
        out bool wasRejectedBySizeLimit) {
        wasRejectedBySizeLimit = false;
        if (recordType == 0xF01A || recordType == 0xF01B || recordType == 0xF01C) {
            return ExtractMetafile(source, offset, count, recordType, recordInstance,
                maximumDecodedImageBytes, out wasRejectedBySizeLimit);
        }

        int imagePrefixLength = GetRasterPrefixLength(recordType, recordInstance);
        if (imagePrefixLength < 0 || imagePrefixLength > count
            || !HasRasterSignature(source, offset + imagePrefixLength,
                count - imagePrefixLength, recordType)) {
            imagePrefixLength = FindRasterSignatureOffset(source, offset, count, recordType);
        }
        if (imagePrefixLength < 0 || imagePrefixLength > count) return Array.Empty<byte>();
        int imageLength = count - imagePrefixLength;
        if (imageLength <= 0) return Array.Empty<byte>();
        if (imageLength > maximumDecodedImageBytes) {
            wasRejectedBySizeLimit = true;
            return Array.Empty<byte>();
        }
        int imageOffset = checked(offset + imagePrefixLength);
        if (recordType == 0xF01F) {
            return CreateBitmapFile(source, imageOffset, imageLength,
                maximumDecodedImageBytes, out wasRejectedBySizeLimit);
        }
        var bytes = new byte[imageLength];
        Buffer.BlockCopy(source, imageOffset, bytes, 0, bytes.Length);
        return bytes;
    }

    private static byte[] ExtractMetafile(byte[] source, int offset, int count, ushort recordType,
        ushort recordInstance, int maximumDecodedImageBytes,
        out bool wasRejectedBySizeLimit) {
        wasRejectedBySizeLimit = false;
        if (recordType == 0xF01C) return Array.Empty<byte>();
        int uidLength = GetMetafileUidLength(recordType, recordInstance);
        if (uidLength < 0 || count < uidLength + 34) return Array.Empty<byte>();
        int headerOffset = checked(offset + uidLength);
        uint uncompressedSize = ReadUInt32(source, headerOffset);
        uint storedSize = ReadUInt32(source, headerOffset + 28);
        byte compression = source[headerOffset + 32];
        int dataOffset = checked(headerOffset + 34);
        int available = Math.Max(0, offset + count - dataOffset);
        int storedLength = storedSize > int.MaxValue
            ? available
            : Math.Min(available, unchecked((int)storedSize));
        if (uncompressedSize > unchecked((uint)maximumDecodedImageBytes)) {
            wasRejectedBySizeLimit = true;
            return Array.Empty<byte>();
        }
        if (storedLength <= 0) {
            return Array.Empty<byte>();
        }
        if (compression == 0xFE) {
            if (storedLength > maximumDecodedImageBytes) {
                wasRejectedBySizeLimit = true;
                return Array.Empty<byte>();
            }
            var bytes = new byte[storedLength];
            Buffer.BlockCopy(source, dataOffset, bytes, 0, bytes.Length);
            return bytes;
        }
        if (compression != 0x00 || storedLength < 6) return Array.Empty<byte>();

        try {
            using var input = new MemoryStream(source, dataOffset + 2, storedLength - 6, writable: false);
            using var inflater = new DeflateStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream(uncompressedSize <= int.MaxValue
                ? unchecked((int)uncompressedSize)
                : 0);
            var buffer = new byte[8192];
            int total = 0;
            int read;
            while ((read = inflater.Read(buffer, 0, buffer.Length)) > 0) {
                total = checked(total + read);
                if (total > maximumDecodedImageBytes) {
                    wasRejectedBySizeLimit = true;
                    return Array.Empty<byte>();
                }
                output.Write(buffer, 0, read);
            }
            return uncompressedSize == 0 || output.Length == uncompressedSize
                ? output.ToArray()
                : Array.Empty<byte>();
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is IOException
                                            || exception is OverflowException) {
            return Array.Empty<byte>();
        }
    }

    private static byte[] CreateBitmapFile(byte[] source, int offset, int count,
        int maximumDecodedImageBytes, out bool wasRejectedBySizeLimit) {
        wasRejectedBySizeLimit = false;
        if (count < 12) return Array.Empty<byte>();
        if ((long)count + 14L > maximumDecodedImageBytes) {
            wasRejectedBySizeLimit = true;
            return Array.Empty<byte>();
        }
        uint dibHeaderSize = ReadUInt32(source, offset);
        if (dibHeaderSize < 12 || dibHeaderSize > unchecked((uint)count)) return Array.Empty<byte>();
        long pixelOffset = dibHeaderSize;
        if (dibHeaderSize == 12) {
            ushort bitsPerPixel = ReadUInt16(source, offset + 10);
            if (bitsPerPixel <= 8) pixelOffset += 3L * (1L << bitsPerPixel);
        } else if (dibHeaderSize >= 40 && count >= 40) {
            ushort bitsPerPixel = ReadUInt16(source, offset + 14);
            uint compression = ReadUInt32(source, offset + 16);
            uint colorCount = ReadUInt32(source, offset + 32);
            if (dibHeaderSize == 40 && compression == 3) pixelOffset += 12;
            if (dibHeaderSize == 40 && compression == 6) pixelOffset += 16;
            if (colorCount == 0 && bitsPerPixel <= 8) colorCount = unchecked((uint)(1 << bitsPerPixel));
            pixelOffset += 4L * colorCount;
        }
        if (pixelOffset < 0 || pixelOffset > count) return Array.Empty<byte>();

        var result = new byte[checked(count + 14)];
        result[0] = 0x42;
        result[1] = 0x4D;
        WriteUInt32(result, 2, unchecked((uint)result.Length));
        WriteUInt32(result, 10, checked((uint)(14L + pixelOffset)));
        Buffer.BlockCopy(source, offset, result, 14, count);
        return result;
    }

    private static int GetRasterPrefixLength(ushort recordType, ushort recordInstance) {
        if (recordType == 0xF01D || recordType == 0xF02A) {
            return recordInstance switch {
                0x046A or 0x06E2 => 17,
                0x046B or 0x06E3 => 33,
                _ => -1
            };
        }
        if (recordType == 0xF01E) return recordInstance switch { 0x06E0 => 17, 0x06E1 => 33, _ => -1 };
        if (recordType == 0xF01F) return recordInstance switch { 0x07A8 => 17, 0x07A9 => 33, _ => -1 };
        if (recordType == 0xF029) return recordInstance switch { 0x06E4 => 17, 0x06E5 => 33, _ => -1 };
        return -1;
    }

    private static int GetMetafileUidLength(ushort recordType, ushort recordInstance) => recordType switch {
        0xF01A => recordInstance switch { 0x03D4 => 16, 0x03D5 => 32, _ => -1 },
        0xF01B => recordInstance switch { 0x0216 => 16, 0x0217 => 32, _ => -1 },
        0xF01C => recordInstance switch { 0x0542 => 16, 0x0543 => 32, _ => -1 },
        _ => -1
    };

    private static int FindRasterSignatureOffset(byte[] source, int offset, int count, ushort recordType) {
        int maximum = Math.Min(count, 64);
        for (int index = 0; index < maximum; index++) {
            if (HasRasterSignature(source, offset + index, maximum - index, recordType)) return index;
        }
        return -1;
    }

    private static bool HasRasterSignature(byte[] source, int offset, int count, ushort recordType) {
        if (count < 0 || offset < 0 || offset > source.Length - count) return false;
        if (recordType == 0xF01F) return count >= 12;
        if (recordType == 0xF01D || recordType == 0xF02A) {
            return count >= 2 && source[offset] == 0xFF && source[offset + 1] == 0xD8;
        }
        if (recordType == 0xF01E) {
            return count >= 4 && source[offset] == 0x89 && source[offset + 1] == 0x50
                && source[offset + 2] == 0x4E && source[offset + 3] == 0x47;
        }
        if (recordType == 0xF029) {
            return count >= 4 && ((source[offset] == 0x49 && source[offset + 1] == 0x49
                                  && source[offset + 2] == 0x2A && source[offset + 3] == 0x00)
                                 || (source[offset] == 0x4D && source[offset + 1] == 0x4D
                                     && source[offset + 2] == 0x00 && source[offset + 3] == 0x2A));
        }
        return false;
    }

    private static string? ReadName(byte[] source, int offset, int boundary, byte count) {
        if (count == 0 || offset >= boundary) return null;
        int available = Math.Min(count, boundary - offset);
        int evenLength = available - available % 2;
        if (evenLength == 0) return null;
        string value = Encoding.Unicode.GetString(source, offset, evenLength).TrimEnd('\0');
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static bool IsBlipRecordType(ushort recordType) => recordType == 0xF01A
        || recordType == 0xF01B || recordType == 0xF01C || recordType == 0xF01D
        || recordType == 0xF01E || recordType == 0xF01F || recordType == 0xF029
        || recordType == 0xF02A;

    private static string ComputeSha256(byte[] source, int offset, int count) {
        using SHA256 sha256 = SHA256.Create();
        return ToHexString(sha256.ComputeHash(source, offset, count), 0, 32);
    }

    private static string ToHexString(byte[] source, int offset, int count) {
        var builder = new StringBuilder(checked(count * 2));
        for (int index = 0; index < count; index++) {
            builder.Append(source[offset + index].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }
        return builder.ToString();
    }

    private static ushort ReadUInt16(byte[] source, int offset) => unchecked((ushort)(
        source[offset] | source[offset + 1] << 8));

    private static uint ReadUInt32(byte[] source, int offset) => unchecked((uint)(
        source[offset] | source[offset + 1] << 8 | source[offset + 2] << 16 | source[offset + 3] << 24));

    private static void WriteUInt32(byte[] target, int offset, uint value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
        target[offset + 2] = unchecked((byte)(value >> 16));
        target[offset + 3] = unchecked((byte)(value >> 24));
    }

    private readonly struct BlipReadResult {
        internal BlipReadResult(OfficeArtBlipStorage storage, byte recordVersion,
            ushort recordInstance, ushort recordType, uint payloadLength, int payloadAvailableLength,
            string? payloadSha256, byte[] imageBytes,
            bool wasImageRejectedBySizeLimit) {
            HasRecord = true;
            Storage = storage;
            RecordVersion = recordVersion;
            RecordInstance = recordInstance;
            RecordType = recordType;
            PayloadLength = payloadLength;
            PayloadAvailableLength = payloadAvailableLength;
            PayloadSha256 = payloadSha256;
            ImageBytes = imageBytes;
            WasImageRejectedBySizeLimit = wasImageRejectedBySizeLimit;
        }

        internal bool HasRecord { get; }
        internal OfficeArtBlipStorage Storage { get; }
        internal byte? RecordVersion { get; }
        internal ushort? RecordInstance { get; }
        internal ushort? RecordType { get; }
        internal uint? PayloadLength { get; }
        internal int? PayloadAvailableLength { get; }
        internal string? PayloadSha256 { get; }
        internal byte[]? ImageBytes { get; }
        internal bool WasImageRejectedBySizeLimit { get; }
    }

    internal readonly struct OfficeArtBlipRecordData {
        private readonly byte[] _imageBytes;

        internal OfficeArtBlipRecordData(byte recordVersion,
            ushort recordInstance, ushort recordType, uint payloadLength,
            int payloadAvailableLength, string? payloadSha256,
            string? contentType, byte[] imageBytes,
            bool wasImageRejectedBySizeLimit) {
            RecordVersion = recordVersion;
            RecordInstance = recordInstance;
            RecordType = recordType;
            PayloadLength = payloadLength;
            PayloadAvailableLength = payloadAvailableLength;
            PayloadSha256 = payloadSha256;
            ContentType = contentType;
            _imageBytes = imageBytes == null
                ? Array.Empty<byte>()
                : (byte[])imageBytes.Clone();
            WasImageRejectedBySizeLimit = wasImageRejectedBySizeLimit;
        }

        internal byte RecordVersion { get; }
        internal ushort RecordInstance { get; }
        internal ushort RecordType { get; }
        internal uint PayloadLength { get; }
        internal int PayloadAvailableLength { get; }
        internal string? PayloadSha256 { get; }
        internal string? ContentType { get; }
        internal bool WasImageRejectedBySizeLimit { get; }
        internal byte[] ImageBytes => _imageBytes == null
            ? Array.Empty<byte>()
            : (byte[])_imageBytes.Clone();
        internal bool IsPayloadTruncated => PayloadLength
            > unchecked((uint)PayloadAvailableLength);
    }
}
