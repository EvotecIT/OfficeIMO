using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Reads the fixed and complex portions of an OfficeArt FOPT property table.</summary>
public static class OfficeArtPropertyTableReader {
    /// <summary>
    /// Reads up to <paramref name="propertyCount"/> entries from a bounded FOPT payload.
    /// A truncated fixed table produces an empty result; truncated complex data is returned at its
    /// available length so callers can preserve and diagnose malformed input safely.
    /// </summary>
    public static IReadOnlyList<OfficeArtProperty> Read(byte[] payload, int offset, int length,
        ushort propertyCount) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        if (offset < 0 || offset > payload.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        if (length < 0 || length > payload.Length - offset) throw new ArgumentOutOfRangeException(nameof(length));
        if (propertyCount == 0) return Array.Empty<OfficeArtProperty>();

        int fixedLength = checked(propertyCount * 6);
        if (fixedLength > length) return Array.Empty<OfficeArtProperty>();

        int endOffset = checked(offset + length);
        int complexOffset = checked(offset + fixedLength);
        ulong[] followingLengths = FollowingComplexLengths(payload, offset, propertyCount, complexOffset, endOffset);
        var properties = new List<OfficeArtProperty>(propertyCount);
        for (int index = 0; index < propertyCount; index++) {
            int propertyOffset = checked(offset + index * 6);
            ushort rawOperationId = ReadUInt16(payload, propertyOffset);
            uint value = ReadUInt32(payload, propertyOffset + 2);
            bool isComplex = (rawOperationId & 0x8000) != 0;
            int? availableLength = null;
            byte[]? complexData = null;
            string? complexText = null;
            if (isComplex) {
                ushort propertyId = checked((ushort)(rawOperationId & 0x3fff));
                availableLength = GetAvailableComplexDataLength(payload, propertyId, value, complexOffset, endOffset,
                    followingLengths[index]);
                if (availableLength.Value > 0) {
                    complexData = new byte[availableLength.Value];
                    Buffer.BlockCopy(payload, complexOffset, complexData, 0, complexData.Length);
                    complexText = TryReadComplexText(complexData, propertyId);
                }
                complexOffset = checked(complexOffset + availableLength.Value);
            }

            properties.Add(new OfficeArtProperty(index, rawOperationId, value, availableLength,
                complexText, complexData));
        }
        return properties;
    }

    /// <summary>Reads a complete byte array as a bounded FOPT payload.</summary>
    public static IReadOnlyList<OfficeArtProperty> Read(byte[] payload, ushort propertyCount) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        return Read(payload, 0, payload.Length, propertyCount);
    }

    private static ushort ReadUInt16(byte[] payload, int offset) => unchecked((ushort)(
        payload[offset] | payload[offset + 1] << 8));

    private static uint ReadUInt32(byte[] payload, int offset) => unchecked((uint)(
        payload[offset]
        | payload[offset + 1] << 8
        | payload[offset + 2] << 16
        | payload[offset + 3] << 24));

    private static ulong[] FollowingComplexLengths(byte[] payload, int offset, int count, int complexStart, int end) {
        var following = new ulong[count];
        ulong required = 0;
        int boundary = end;
        for (int index = count - 1; index >= 0; index--) {
            following[index] = required;
            int propertyOffset = offset + index * 6;
            ushort operation = ReadUInt16(payload, propertyOffset);
            if ((operation & 0x8000) == 0) continue;
            uint declared = ReadUInt32(payload, propertyOffset + 2), length = declared;
            ushort id = (ushort)(operation & 0x3fff);
            int remaining = boundary - complexStart;
            if (declared > 0 && declared <= remaining) {
                int start = boundary - (int)declared;
                // Prefer a recognizable header-inclusive suffix. Otherwise a
                // complete header-excluded suffix reserves its extra six bytes.
                bool includesHeader = declared >= 6 && ExpectedComplexDataLength(id, declared - 6,
                    payload, start, (int)declared) == declared;
                if (!includesHeader && declared <= remaining - 6 && ExpectedComplexDataLength(id, declared,
                        payload, start - 6, (int)declared + 6) == declared + 6)
                    length += 6;
            }
            // Zero-length arrays can be absent; a reverse scan cannot establish
            // their header without risking bytes owned by the preceding payload.
            required += length;
            boundary = required <= (ulong)(end - complexStart) ? end - (int)required : complexStart;
        }
        return following;
    }

    private static int GetAvailableComplexDataLength(byte[] payload, ushort propertyId,
        uint declaredLength, int offset, int endOffset, ulong followingDeclaredLength) {
        if (declaredLength > int.MaxValue || offset >= endOffset) return 0;
        uint length = ExpectedComplexDataLength(propertyId, declaredLength, payload, offset, endOffset - offset);
        int remaining = endOffset - offset;
        if (length > declaredLength) {
            // Header recovery must preserve following complex properties even
            // when a nonempty array is truncated. Retain only its available prefix.
            uint available = (uint)((ulong)remaining - Math.Min((ulong)remaining, followingDeclaredLength));
            // A zero-length array can be genuinely absent, so do not consume a
            // partial or absent header as if it were its own complex payload.
            if (declaredLength == 0 && available < length) return 0;
            return (int)Math.Min(length, available);
        }
        return (int)Math.Min(length, (uint)remaining);
    }

    internal static uint ExpectedComplexDataLength(ushort propertyId, uint declaredLength,
        byte[]? payload, int offset, int remaining) {
        // These native geometry/color arrays can declare only element bytes.
        // Recognize their exact header layout without guessing for arbitrary complex properties.
        if (payload != null && remaining >= 6 && propertyId is 0x0145 or 0x0146 or 0x0156 or 0x0197) {
            int count = ReadUInt16(payload, offset), allocated = ReadUInt16(payload, offset + 2);
            int size = ReadUInt16(payload, offset + 4);
            bool truncatedGuide = propertyId == 0x0156 && size == 0xFFF0;
            if ((propertyId == 0x0145 || truncatedGuide) && size == 0xFFF0) size = 4;
            bool knownSize = propertyId switch { 0x0145 => size is 4 or 8, 0x0146 => size == 2,
                0x0156 => size == 8 || truncatedGuide, _ => size == 8 };
            if (knownSize && count <= allocated && count * size == declaredLength) return declaredLength + 6;
        }
        return declaredLength;
    }

    private static string? TryReadComplexText(byte[] data, ushort propertyId) {
        if (propertyId is not 0x0380 and not 0x0381 || data.Length < 2) return null;
        int evenLength = data.Length - data.Length % 2;
        string value = Encoding.Unicode.GetString(data, 0, evenLength).TrimEnd('\0');
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }
}
