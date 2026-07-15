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
                availableLength = GetAvailableComplexDataLength(value, complexOffset, endOffset);
                if (availableLength.Value > 0) {
                    complexData = new byte[availableLength.Value];
                    Buffer.BlockCopy(payload, complexOffset, complexData, 0, complexData.Length);
                    ushort propertyId = checked((ushort)(rawOperationId & 0x3fff));
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

    private static int GetAvailableComplexDataLength(uint declaredLength, int offset, int endOffset) {
        if (declaredLength > int.MaxValue || offset >= endOffset) return 0;
        return Math.Min((int)declaredLength, endOffset - offset);
    }

    private static string? TryReadComplexText(byte[] data, ushort propertyId) {
        if (propertyId is not 0x0380 and not 0x0381 || data.Length < 2) return null;
        int evenLength = data.Length - data.Length % 2;
        string value = Encoding.Unicode.GetString(data, 0, evenLength).TrimEnd('\0');
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }
}
