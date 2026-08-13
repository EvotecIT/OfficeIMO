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
    private static readonly byte[] ClaimUuid = {
        0x63, 0x32, 0x63, 0x6C, 0x00, 0x11, 0x00, 0x10,
        0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
    };
    private static readonly byte[] ClaimSignatureUuid = {
        0x63, 0x32, 0x63, 0x73, 0x00, 0x11, 0x00, 0x10,
        0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
    };

    internal static bool IsValid(
        byte[] data,
        int offset,
        int availableLength,
        long maximumBytes,
        int maximumEntries,
        out int storeLength) {
        storeLength = 0;
        int visitedBoxes = 0;
        if (offset < 0 || availableLength < 0 || offset > data.Length - availableLength || availableLength < 38) return false;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength > (ulong)maximumBytes || declaredLength != (ulong)availableLength || declaredLength > int.MaxValue) {
            return false;
        }

        int totalLength = (int)declaredLength;
        int childOffset = offset + headerLength;
        int childAvailable = totalLength - headerLength;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
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
            if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
            int remaining = storeEnd - nextChildOffset;
            if (!TryReadBox(data, nextChildOffset, remaining, out _, out ulong nextChildLength, out string nextChildType) ||
                nextChildLength > int.MaxValue) {
                return false;
            }
            if (nextChildType == "jumb" && IsManifestSuperbox(
                data, nextChildOffset, (int)nextChildLength, ref visitedBoxes, maximumEntries)) {
                hasManifest = true;
            }
            nextChildOffset += (int)nextChildLength;
        }
        if (nextChildOffset != storeEnd || !hasManifest) return false;

        storeLength = totalLength;
        return true;
    }

    private static bool IsManifestSuperbox(
        byte[] data,
        int offset,
        int availableLength,
        ref int visitedBoxes,
        int maximumEntries) {
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength != (ulong)availableLength) return false;
        int descriptionOffset = offset + headerLength;
        int descriptionAvailable = availableLength - headerLength;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        if (!TryReadBox(data, descriptionOffset, descriptionAvailable, out int descriptionHeaderLength,
            out ulong descriptionLength, out string descriptionType) || descriptionType != "jumd" ||
            descriptionLength < (ulong)(descriptionHeaderLength + 18)) return false;
        int payloadOffset = descriptionOffset + descriptionHeaderLength;
        if (!Matches(data, payloadOffset, StandardManifestUuid) && !Matches(data, payloadOffset, UpdateManifestUuid)) return false;
        int togglesOffset = payloadOffset + StandardManifestUuid.Length;
        if ((data[togglesOffset] & 0x02) == 0) return false;
        int labelOffset = togglesOffset + 1;
        int descriptionEnd = descriptionOffset + (int)descriptionLength;
        if (labelOffset >= descriptionEnd || data[labelOffset] == 0 ||
            Array.IndexOf(data, (byte)0, labelOffset, descriptionEnd - labelOffset) < 0) return false;

        int manifestEnd = offset + availableLength;
        int cursor = descriptionEnd;
        int claimCount = 0;
        int claimSignatureCount = 0;
        while (cursor < manifestEnd) {
            if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
            int remaining = manifestEnd - cursor;
            if (!TryReadBox(data, cursor, remaining, out _, out ulong childLength, out string childType) ||
                childLength > int.MaxValue) return false;
            if (childType == "jumb" && HasDescriptionUuid(data, cursor, (int)childLength, ClaimUuid)) {
                claimCount++;
                if (!IsClaimSuperbox(data, cursor, (int)childLength, ref visitedBoxes, maximumEntries)) return false;
            } else if (childType == "jumb" && HasDescriptionUuid(data, cursor, (int)childLength, ClaimSignatureUuid)) {
                claimSignatureCount++;
                if (!IsClaimSignatureSuperbox(data, cursor, (int)childLength, ref visitedBoxes, maximumEntries)) return false;
            }
            cursor += (int)childLength;
        }
        return cursor == manifestEnd && claimCount == 1 && claimSignatureCount == 1;
    }

    private static bool HasDescriptionUuid(byte[] data, int offset, int availableLength, byte[] uuid) {
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength != (ulong)availableLength) return false;
        int descriptionOffset = offset + headerLength;
        return TryReadBox(data, descriptionOffset, availableLength - headerLength, out int descriptionHeaderLength,
            out ulong descriptionLength, out string descriptionType) && descriptionType == "jumd" &&
            descriptionLength >= (ulong)(descriptionHeaderLength + uuid.Length) &&
            Matches(data, descriptionOffset + descriptionHeaderLength, uuid);
    }

    private static bool IsClaimSuperbox(
        byte[] data,
        int offset,
        int availableLength,
        ref int visitedBoxes,
        int maximumEntries) {
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength != (ulong)availableLength) return false;
        int descriptionOffset = offset + headerLength;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        if (!TryReadBox(data, descriptionOffset, availableLength - headerLength, out int descriptionHeaderLength,
            out ulong descriptionLength, out string descriptionType) || descriptionType != "jumd" ||
            descriptionLength < (ulong)(descriptionHeaderLength + ClaimUuid.Length + 2)) return false;
        int payloadOffset = descriptionOffset + descriptionHeaderLength;
        if (!Matches(data, payloadOffset, ClaimUuid)) return false;
        int togglesOffset = payloadOffset + ClaimUuid.Length;
        if ((data[togglesOffset] & 0x02) == 0) return false;
        int labelOffset = togglesOffset + 1;
        int descriptionEnd = descriptionOffset + (int)descriptionLength;
        int terminator = Array.IndexOf(data, (byte)0, labelOffset, descriptionEnd - labelOffset);
        if (terminator < 0) return false;
        string label = System.Text.Encoding.ASCII.GetString(data, labelOffset, terminator - labelOffset);
        if (label != "c2pa.claim" && label != "c2pa.claim.v2") return false;
        int contentOffset = descriptionEnd;
        int contentAvailable = offset + availableLength - contentOffset;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        return TryReadBox(data, contentOffset, contentAvailable, out int contentHeaderLength,
            out ulong contentLength, out string contentType) && contentType == "cbor" &&
            contentLength > (ulong)contentHeaderLength && contentLength == (ulong)contentAvailable;
    }

    private static bool IsClaimSignatureSuperbox(
        byte[] data,
        int offset,
        int availableLength,
        ref int visitedBoxes,
        int maximumEntries) {
        if (!TryReadBox(data, offset, availableLength, out int headerLength, out ulong declaredLength, out string type) ||
            type != "jumb" || declaredLength != (ulong)availableLength) return false;
        int descriptionOffset = offset + headerLength;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        if (!TryReadBox(data, descriptionOffset, availableLength - headerLength, out int descriptionHeaderLength,
            out ulong descriptionLength, out string descriptionType) || descriptionType != "jumd" ||
            descriptionLength < (ulong)(descriptionHeaderLength + ClaimSignatureUuid.Length + 2)) return false;
        int payloadOffset = descriptionOffset + descriptionHeaderLength;
        if (!Matches(data, payloadOffset, ClaimSignatureUuid)) return false;
        int togglesOffset = payloadOffset + ClaimSignatureUuid.Length;
        if ((data[togglesOffset] & 0x02) == 0) return false;
        int labelOffset = togglesOffset + 1;
        int descriptionEnd = descriptionOffset + (int)descriptionLength;
        int terminator = Array.IndexOf(data, (byte)0, labelOffset, descriptionEnd - labelOffset);
        if (terminator < 0 || System.Text.Encoding.ASCII.GetString(data, labelOffset, terminator - labelOffset) != "c2pa.signature") {
            return false;
        }
        int contentOffset = descriptionEnd;
        int contentAvailable = offset + availableLength - contentOffset;
        if (!TryReserveBox(ref visitedBoxes, maximumEntries)) return false;
        return TryReadBox(data, contentOffset, contentAvailable, out int contentHeaderLength,
            out ulong contentLength, out string contentType) && contentType == "cbor" &&
            contentLength > (ulong)contentHeaderLength && contentLength == (ulong)contentAvailable;
    }

    private static bool TryReserveBox(ref int visitedBoxes, int maximumEntries) {
        if (visitedBoxes >= maximumEntries) return false;
        visitedBoxes++;
        return true;
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
