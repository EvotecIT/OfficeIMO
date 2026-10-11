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
    /// Complete recognizable array prefixes take priority over overlapping suffix interpretations;
    /// recovery of malformed complex-data boundaries is best effort.
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
        ulong[] followingLengths = FollowingComplexLengths(payload, offset, propertyCount, complexOffset, endOffset,
            out bool[] followingArrayBoundaries);
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
                    followingLengths[index], followingArrayBoundaries[index]);
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

    private static ulong[] FollowingComplexLengths(byte[] payload, int offset, int count, int complexStart, int end,
        out bool[] arrayBoundaries) {
        var following = new ulong[count];
        arrayBoundaries = new bool[count];
        // A recognizable empty suffix can also be the tail of a complete earlier
        // array. Keep established complete prefix boundaries before reserving it.
        var prefixStarts = new int[count];
        int prefix = complexStart;
        for (int index = 0; index < count; index++) {
            prefixStarts[index] = prefix;
            int propertyOffset = offset + index * 6;
            ushort operation = ReadUInt16(payload, propertyOffset);
            if (prefix < 0 || (operation & 0x8000) == 0) continue;
            uint declared = ReadUInt32(payload, propertyOffset + 2);
            uint physical = ExpectedComplexDataLength((ushort)(operation & 0x3fff), declared,
                payload, prefix, end - prefix);
            prefix = physical <= end - prefix ? prefix + (int)physical : -1;
        }
        ulong required = 0;
        int boundary = end;
        bool hasArrayBoundary = false;
        for (int index = count - 1; index >= 0; index--) {
            following[index] = required;
            arrayBoundaries[index] = hasArrayBoundary;
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
                if (includesHeader) hasArrayBoundary = true;
                else if (declared <= remaining - 6 && ExpectedComplexDataLength(id, declared,
                        payload, start - 6, (int)declared + 6) == declared + 6) {
                    length += 6;
                    hasArrayBoundary = true;
                }
            }
            if (declared == 0 && remaining >= 6 &&
                (prefixStarts[index] < 0 || prefixStarts[index] <= boundary - 6) &&
                ExpectedComplexDataLength(id, 0, payload, boundary - 6, 6) == 6) {
                length = 6;
                hasArrayBoundary = true;
            }
            required += length;
            if (required > (ulong)(end - complexStart)) hasArrayBoundary = false;
            boundary = required <= (ulong)(end - complexStart) ? end - (int)required : complexStart;
        }
        return following;
    }

    private static int GetAvailableComplexDataLength(byte[] payload, ushort propertyId,
        uint declaredLength, int offset, int endOffset, ulong followingDeclaredLength, bool followingArrayBoundary) {
        if (offset >= endOffset) return 0;
        uint length = ExpectedComplexDataLength(propertyId, declaredLength, payload, offset, endOffset - offset);
        int remaining = endOffset - offset;
        bool arrayHeader = length > declaredLength || (declaredLength >= 6 &&
            ExpectedComplexDataLength(propertyId, declaredLength - 6, payload, offset, remaining) == declaredLength);
        // A physically complete recognized prefix takes priority over suffix
        // guesses. Array coordinates can resemble headers of absent properties.
        if (arrayHeader && declaredLength != 0 && length <= (uint)remaining) return (int)length;
        if (arrayHeader || followingArrayBoundary) {
            // An oversized/truncated suffix has no established boundary. It must
            // not erase a complete, recognizable array at the current position.
            if (followingDeclaredLength > (ulong)remaining)
                return (int)Math.Min(length, (uint)remaining);
            // Recognizable suffix arrays also establish a boundary for truncated
            // ordinary payloads. Keep that boundary and the declared trailing data.
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
            bool compact = size == 0xFFF0 && propertyId is 0x0145 or 0x0156 or 0x0197;
            if (compact) size = 4;
            bool knownSize = propertyId switch { 0x0145 => size is 4 or 8, 0x0146 => size == 2,
                0x0156 or 0x0197 => size == 8 || compact, _ => false };
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
