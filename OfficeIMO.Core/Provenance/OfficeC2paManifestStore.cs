using System;

namespace OfficeIMO.Provenance;

internal static class OfficeC2paManifestStore {
    private static readonly byte[] ManifestStoreUuid = {
        0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10,
        0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
    };
    private static readonly byte[] StandardManifestUuid = {
        0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10,
        0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
    };
    private static readonly byte[] UpdateManifestUuid = {
        0x63, 0x32, 0x75, 0x6D, 0x00, 0x11, 0x00, 0x10,
        0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
    };

    internal static bool IsValid(byte[] data, int offset, int availableLength, long maximumBytes, out int storeLength) {
        storeLength = 0;
        if (offset < 0 || availableLength < 0 || offset > data.Length - availableLength || availableLength < 38) return false;
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength > (ulong)maximumBytes || declaredLength != (ulong)availableLength || declaredLength > int.MaxValue) {
            return false;
        }

        int totalLength = (int)declaredLength;
        int childOffset = offset + headerLength;
        int childAvailable = totalLength - headerLength;
        if (!TryReadBox(data, childOffset, childAvailable, out int childHeaderLength, out ulong childLength, out string childType) ||
            childType != "jumd" || childLength < (ulong)(childHeaderLength + 22) || childLength > (ulong)childAvailable) {
            return false;
        }

        int payloadOffset = childOffset + childHeaderLength;
        if (!Matches(data, payloadOffset, ManifestStoreUuid)) return false;
        int togglesOffset = payloadOffset + ManifestStoreUuid.Length;
        byte toggles = data[togglesOffset];
        if ((toggles & 0x02) == 0) return false;
        int labelOffset = togglesOffset + 1;
        if (!OfficeProvenanceBinary.MatchesAscii(data, labelOffset, "c2pa") ||
            labelOffset + 4 >= childOffset + (int)childLength || data[labelOffset + 4] != 0) {
            return false;
        }

        int storeEnd = offset + totalLength;
        int nextChildOffset = childOffset + (int)childLength;
        bool hasManifest = false;
        while (nextChildOffset < storeEnd) {
            int remaining = storeEnd - nextChildOffset;
            if (!TryReadBox(data, nextChildOffset, remaining, out _, out ulong nextChildLength, out string nextChildType) ||
                nextChildLength > int.MaxValue) {
                return false;
            }
            if (nextChildType == "jumb" && IsManifestSuperbox(data, nextChildOffset, (int)nextChildLength)) {
                hasManifest = true;
            }
            nextChildOffset += (int)nextChildLength;
        }
        if (nextChildOffset != storeEnd || !hasManifest) return false;

        storeLength = totalLength;
        return true;
    }

    private static bool IsManifestSuperbox(byte[] data, int offset, int availableLength) {
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength != (ulong)availableLength) return false;
        int descriptionOffset = offset + headerLength;
        int descriptionAvailable = availableLength - headerLength;
        if (!TryReadBox(data, descriptionOffset, descriptionAvailable, out int descriptionHeaderLength,
            out ulong descriptionLength, out string descriptionType) || descriptionType != "jumd" ||
            descriptionLength < (ulong)(descriptionHeaderLength + 18)) return false;
        int payloadOffset = descriptionOffset + descriptionHeaderLength;
        if (!Matches(data, payloadOffset, StandardManifestUuid) && !Matches(data, payloadOffset, UpdateManifestUuid)) return false;
        int togglesOffset = payloadOffset + StandardManifestUuid.Length;
        if ((data[togglesOffset] & 0x02) == 0) return false;
        int labelOffset = togglesOffset + 1;
        int descriptionEnd = descriptionOffset + (int)descriptionLength;
        return labelOffset < descriptionEnd && data[labelOffset] != 0 &&
            Array.IndexOf(data, (byte)0, labelOffset, descriptionEnd - labelOffset) >= 0;
    }

    private static bool TryReadBox(
        byte[] data,
        int offset,
        int availableLength,
        out int headerLength,
        out ulong declaredLength,
        out string type) {
        headerLength = 0;
        declaredLength = 0;
        type = string.Empty;
        if (availableLength < 8 || offset < 0 || offset > data.Length - availableLength) return false;
        uint length32 = OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: false);
        type = System.Text.Encoding.ASCII.GetString(data, offset + 4, 4);
        if (length32 == 1) {
            if (availableLength < 16) return false;
            declaredLength = OfficeProvenanceBinary.ReadUInt64(data, offset + 8, littleEndian: false);
            headerLength = 16;
        } else if (length32 == 0) {
            declaredLength = (ulong)availableLength;
            headerLength = 8;
        } else {
            declaredLength = length32;
            headerLength = 8;
        }
        return declaredLength >= (ulong)headerLength && declaredLength <= (ulong)availableLength;
    }

    private static bool Matches(byte[] data, int offset, byte[] expected) {
        if (offset < 0 || expected.Length > data.Length - offset) return false;
        for (int index = 0; index < expected.Length; index++) {
            if (data[offset + index] != expected[index]) return false;
        }
        return true;
    }
}
