using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceJpeg {
    private static readonly byte[] XmpHeader = System.Text.Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes, out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        using var output = new MemoryStream(data.Length);
        reserialized = Walk(data, options.Limits, context: null, output, options, changes);
        return output.ToArray();
    }

    private static bool Walk(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        Stream? output,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes) {
        bool reserialized = false;
        int searchOffset = 0;
        int imageIndex = 0;
        while (searchOffset < data.Length) {
            int imageStart = FindNextStart(data, searchOffset);
            if (imageStart < 0) {
                if (output != null && searchOffset < data.Length) output.Write(data, searchOffset, data.Length - searchOffset);
                return reserialized;
            }
            if (output != null && searchOffset < imageStart) output.Write(data, searchOffset, imageStart - searchOffset);
            output?.Write(data, imageStart, 2);
            int offset = imageStart + 2;
            while (offset < data.Length) {
                int segmentStart = offset;
                if (!TryReadMarker(data, segmentStart, out byte marker, out int payloadOffset, out int payloadLength, out int segmentEnd)) {
                    throw new InvalidDataException("JPEG contains an invalid marker sequence.");
                }
                if (marker == 0xD9) {
                    output?.Write(data, segmentStart, segmentEnd - segmentStart);
                    searchOffset = segmentEnd;
                    imageIndex++;
                    break;
                }
                if (marker == 0xDA) {
                    int imageEnd = FindEnd(data, segmentStart);
                    output?.Write(data, segmentStart, imageEnd - segmentStart);
                    searchOffset = imageEnd;
                    imageIndex++;
                    break;
                }

                if (marker == 0xE1 && HasXmpHeader(data, payloadOffset, payloadLength)) {
                    int packetOffset = payloadOffset + XmpHeader.Length;
                    int packetLength = payloadLength - XmpHeader.Length;
                    byte[] packet = new byte[packetLength];
                    Buffer.BlockCopy(data, packetOffset, packet, 0, packetLength);
                    string location = $"JPEG[{imageIndex}]/APP1-XMP@{segmentStart}";
                    if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
                    if (output != null && removalOptions != null && changes != null &&
                        OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                        byte[] payload = new byte[XmpHeader.Length + cleaned.Length];
                        Buffer.BlockCopy(XmpHeader, 0, payload, 0, XmpHeader.Length);
                        Buffer.BlockCopy(cleaned, 0, payload, XmpHeader.Length, cleaned.Length);
                        WriteSegment(output, marker, payload);
                        reserialized = true;
                    } else {
                        output?.Write(data, segmentStart, segmentEnd - segmentStart);
                    }
                    offset = segmentEnd;
                    continue;
                }

                if (marker == 0xEB && TryGetSequenceEnd(data, segmentStart, payloadOffset, payloadLength, options, out int sequenceEnd, out int manifestLength)) {
                    string location = $"JPEG[{imageIndex}]/APP11@{segmentStart}";
                    context?.Add(new OfficeProvenanceEvidence(
                        OfficeProvenanceCarrierKind.C2paManifest,
                        location,
                        isStructurallyValid: true,
                        payloadLength: manifestLength));
                    if (output != null && removalOptions != null && changes != null && removalOptions.RemoveC2paManifests) {
                        changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, sequenceEnd - segmentStart));
                    } else {
                        output?.Write(data, segmentStart, sequenceEnd - segmentStart);
                    }
                    offset = sequenceEnd;
                    continue;
                }

                output?.Write(data, segmentStart, segmentEnd - segmentStart);
                offset = segmentEnd;
            }
            if (offset >= data.Length && searchOffset <= imageStart) throw new InvalidDataException("JPEG does not contain an end marker.");
        }
        return reserialized;
    }

    private static bool HasXmpHeader(byte[] data, int payloadOffset, int payloadLength) {
        if (payloadLength < XmpHeader.Length) return false;
        for (int index = 0; index < XmpHeader.Length; index++) if (data[payloadOffset + index] != XmpHeader[index]) return false;
        return true;
    }

    private static void WriteSegment(Stream output, byte marker, byte[] payload) {
        if (payload.Length > ushort.MaxValue - 2) throw new InvalidDataException("JPEG XMP packet exceeds the APP1 segment limit.");
        output.WriteByte(0xFF);
        output.WriteByte(marker);
        int length = payload.Length + 2;
        output.WriteByte((byte)(length >> 8));
        output.WriteByte((byte)length);
        output.Write(payload, 0, payload.Length);
    }

    private static bool TryGetSequenceEnd(
        byte[] data,
        int segmentStart,
        int payloadOffset,
        int payloadLength,
        OfficeProvenanceOptions options,
        out int sequenceEnd,
        out int manifestLength) {
        sequenceEnd = segmentStart;
        manifestLength = 0;
        if (payloadLength < 16 || data[payloadOffset] != 0x4A || data[payloadOffset + 1] != 0x50 ||
            OfficeProvenanceBinary.ReadUInt32(data, payloadOffset + 4, littleEndian: false) != 1) return false;
        int firstFragmentLength = payloadLength - 8;
        if (!OfficeC2paManifestStore.IsValid(data, payloadOffset + 8, firstFragmentLength, options.MaxManifestBytes, out int declaredManifestLength)) {
            // The first APP11 fragment frequently contains only the beginning of a larger JUMBF box.
            if (!TryReadDeclaredJumbfLength(data, payloadOffset + 8, firstFragmentLength, options.MaxManifestBytes, out declaredManifestLength) ||
                !HasC2paDescriptionPrefix(data, payloadOffset + 8, firstFragmentLength)) return false;
        }

        long collected = firstFragmentLength;
        int current = payloadOffset + payloadLength;
        byte instanceHigh = data[payloadOffset + 2];
        byte instanceLow = data[payloadOffset + 3];
        uint expectedSequence = 2;
        while (collected < declaredManifestLength) {
            if (!TryReadMarker(data, current, out byte marker, out int nextPayloadOffset, out int nextPayloadLength, out int nextEnd) ||
                marker != 0xEB || nextPayloadLength < 8 ||
                data[nextPayloadOffset] != 0x4A || data[nextPayloadOffset + 1] != 0x50 ||
                data[nextPayloadOffset + 2] != instanceHigh || data[nextPayloadOffset + 3] != instanceLow ||
                OfficeProvenanceBinary.ReadUInt32(data, nextPayloadOffset + 4, littleEndian: false) != expectedSequence) return false;
            collected += nextPayloadLength - 8L;
            if (collected > declaredManifestLength) return false;
            current = nextEnd;
            expectedSequence++;
        }
        if (collected != declaredManifestLength) return false;
        sequenceEnd = current;
        manifestLength = declaredManifestLength;
        return true;
    }

    private static bool TryReadDeclaredJumbfLength(byte[] data, int offset, int available, long maximum, out int length) {
        length = 0;
        if (available < 8 || !OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "jumb")) return false;
        uint value = OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: false);
        if (value == 0 || value == 1 || value > maximum || value > int.MaxValue) return false;
        length = (int)value;
        return length >= 38;
    }

    private static bool HasC2paDescriptionPrefix(byte[] data, int offset, int available) {
        if (available < 38 || !OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "jumb") ||
            !OfficeProvenanceBinary.MatchesAscii(data, offset + 12, "jumd")) return false;
        byte[] uuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        for (int index = 0; index < uuid.Length; index++) if (data[offset + 16 + index] != uuid[index]) return false;
        return (data[offset + 32] & 0x02) != 0 && OfficeProvenanceBinary.MatchesAscii(data, offset + 33, "c2pa") && data[offset + 37] == 0;
    }

    private static bool TryReadMarker(byte[] data, int start, out byte marker, out int payloadOffset, out int payloadLength, out int end) {
        marker = 0;
        payloadOffset = payloadLength = 0;
        end = start;
        if (start < 0 || start >= data.Length || data[start] != 0xFF) return false;
        int offset = start;
        while (offset < data.Length && data[offset] == 0xFF) offset++;
        if (offset >= data.Length) return false;
        marker = data[offset++];
        if (marker == 0x00) return false;
        if (marker == 0x01 || marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7)) {
            end = offset;
            return true;
        }
        if (data.Length - offset < 2) return false;
        int segmentLength = (data[offset] << 8) | data[offset + 1];
        if (segmentLength < 2 || segmentLength > data.Length - offset) return false;
        payloadOffset = offset + 2;
        payloadLength = segmentLength - 2;
        end = offset + segmentLength;
        return true;
    }

    private static int FindNextStart(byte[] data, int offset) {
        for (int index = Math.Max(0, offset); index + 1 < data.Length; index++) {
            if (data[index] == 0xFF && data[index + 1] == 0xD8) return index;
        }
        return -1;
    }

    private static int FindEnd(byte[] data, int scanStart) {
        if (!TryReadMarker(data, scanStart, out byte marker, out _, out _, out int offset) || marker != 0xDA) {
            throw new InvalidDataException("JPEG scan marker is invalid.");
        }
        while (offset < data.Length - 1) {
            if (data[offset] != 0xFF) { offset++; continue; }
            int markerOffset = offset;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) break;
            byte value = data[offset++];
            if (value == 0x00 || value == 0x01 || (value >= 0xD0 && value <= 0xD7)) continue;
            if (value == 0xD9) return offset;
            if (value == 0xD8) throw new InvalidDataException($"Unexpected JPEG SOI marker in scan data at offset {markerOffset}.");
            if (data.Length - offset < 2) break;
            int length = (data[offset] << 8) | data[offset + 1];
            if (length < 2 || length > data.Length - offset) break;
            offset += length;
        }
        throw new InvalidDataException("JPEG scan does not contain an end marker.");
    }
}
