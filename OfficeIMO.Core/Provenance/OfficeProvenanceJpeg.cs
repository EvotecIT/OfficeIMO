using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceJpeg {

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
        int markerCount = 0;
        while (searchOffset < data.Length) {
            int imageStart = imageIndex == 0
                ? FindNextStart(data, searchOffset)
                : FindNextCompleteStart(data, searchOffset, options.MaxContainerEntries);
            if (imageStart < 0) {
                if (output != null && searchOffset < data.Length) output.Write(data, searchOffset, data.Length - searchOffset);
                SortBySourceOffset(context?.Evidence);
                SortBySourceOffset(changes);
                return reserialized;
            }
            if (output != null && searchOffset < imageStart) output.Write(data, searchOffset, imageStart - searchOffset);
            output?.Write(data, imageStart, 2);
            ReserveMarker(ref markerCount, options.MaxContainerEntries);
            int offset = imageStart + 2;
            OfficeProvenanceJpegXmpResult xmpResult = OfficeProvenanceJpegXmp.ProcessImage(
                data, offset, imageIndex, options, context, removalOptions, changes);
            bool hasDuplicateC2paSequences = CountC2paSequences(data, offset, options, out bool hasImageFrameAndScan) > 1;
            while (offset < data.Length) {
                int segmentStart = offset;
                if (!TryReadMarker(data, segmentStart, out byte marker, out int payloadOffset, out int payloadLength, out int segmentEnd)) {
                    throw new InvalidDataException("JPEG contains an invalid marker sequence.");
                }
                ReserveMarker(ref markerCount, options.MaxContainerEntries);
                if (marker == 0xD8) throw new InvalidDataException("JPEG contains a nested start-of-image marker.");
                if (marker == 0xD9) {
                    output?.Write(data, segmentStart, segmentEnd - segmentStart);
                    searchOffset = segmentEnd;
                    imageIndex++;
                    break;
                }
                if (marker == 0xDA) {
                    int imageEnd = FindEnd(data, segmentStart, ref markerCount, options.MaxContainerEntries);
                    output?.Write(data, segmentStart, imageEnd - segmentStart);
                    searchOffset = imageEnd;
                    imageIndex++;
                    break;
                }

                if (marker == 0xE1 && xmpResult.SegmentStarts.Contains(segmentStart)) {
                    if (output != null && xmpResult.Replacements.TryGetValue(segmentStart, out byte[]? replacement)) {
                        output.Write(replacement, 0, replacement.Length);
                        reserialized = true;
                    } else output?.Write(data, segmentStart, segmentEnd - segmentStart);
                    offset = segmentEnd;
                    continue;
                }

                if (marker == 0xEB && TryGetC2paSequence(data, segmentStart, payloadOffset, payloadLength, options, ref markerCount,
                    out int sequenceEnd, out int manifestLength, out bool structurallyValid)) {
                    structurallyValid &= !hasDuplicateC2paSequences && hasImageFrameAndScan;
                    string location = $"JPEG[{imageIndex}]/APP11@{segmentStart}";
                    context?.Add(new OfficeProvenanceEvidence(
                        OfficeProvenanceCarrierKind.C2paManifest,
                        location,
                        structurallyValid,
                        payloadLength: manifestLength));
                    if (output != null && removalOptions != null && changes != null && removalOptions.RemoveC2paManifests &&
                        (structurallyValid || !removalOptions.RequireStructurallyValidCarrier)) {
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
        SortBySourceOffset(context?.Evidence);
        SortBySourceOffset(changes);
        return reserialized;
    }

    private static int CountC2paSequences(
        byte[] data,
        int offset,
        OfficeProvenanceOptions options,
        out bool hasImageFrameAndScan) {
        int count = 0;
        int markers = 0;
        bool hasValidFrame = false;
        hasImageFrameAndScan = false;
        while (offset < data.Length) {
            if (!TryReadMarker(data, offset, out byte marker, out int payloadOffset, out int payloadLength, out int segmentEnd)) {
                throw new InvalidDataException("JPEG contains an invalid marker sequence.");
            }
            ReserveMarker(ref markers, options.MaxContainerEntries);
            if (marker == 0xD8) throw new InvalidDataException("JPEG contains a nested start-of-image marker.");
            if (marker == 0xD9) return count;
            if (marker == 0xDA) {
                int scanMarkers = markers;
                _ = FindEnd(data, offset, ref scanMarkers, options.MaxContainerEntries);
                hasImageFrameAndScan = hasValidFrame && IsValidStartOfScan(data, payloadOffset, payloadLength);
                return count;
            }
            if (IsStartOfFrame(marker)) {
                hasValidFrame |= IsValidStartOfFrame(data, payloadOffset, payloadLength);
            }
            if (marker == 0xEB && IsC2paSequenceStart(data, payloadOffset, payloadLength)) {
                count++;
                if (TryGetC2paSequence(data, offset, payloadOffset, payloadLength, options, ref markers,
                    out int sequenceEnd, out _, out _)) {
                    offset = sequenceEnd;
                } else {
                    offset = segmentEnd;
                }
            } else if (marker == 0xEB && IsC2paContinuationFragment(data, payloadOffset, payloadLength)) {
                // A continuation that was not consumed by the immediately preceding sequence is
                // competing malformed carrier evidence. Counting it keeps strict mutation from
                // deleting a valid sequence while leaving an orphaned C2PA fragment behind.
                count++;
                offset = segmentEnd;
            } else {
                offset = segmentEnd;
            }
        }
        throw new InvalidDataException("JPEG does not contain an end marker.");
    }

    private static bool IsStartOfFrame(byte marker) => marker is
        0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or
        0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;

    private static bool IsValidStartOfFrame(byte[] data, int payloadOffset, int payloadLength) {
        if (payloadLength < 9) return false;
        int components = data[payloadOffset + 5];
        ushort height = OfficeProvenanceBinary.ReadUInt16(data, payloadOffset + 1, littleEndian: false);
        ushort width = OfficeProvenanceBinary.ReadUInt16(data, payloadOffset + 3, littleEndian: false);
        return components > 0 && payloadLength == 6 + 3 * components && width != 0 && height != 0;
    }

    private static bool IsValidStartOfScan(byte[] data, int payloadOffset, int payloadLength) {
        if (payloadLength < 6) return false;
        int components = data[payloadOffset];
        return components > 0 && payloadLength == 1 + 2 * components + 3;
    }

    private static bool IsC2paSequenceStart(byte[] data, int payloadOffset, int payloadLength) =>
        payloadLength >= 8 && data[payloadOffset] == 0x4A && data[payloadOffset + 1] == 0x50 &&
        OfficeProvenanceBinary.ReadUInt32(data, payloadOffset + 4, littleEndian: false) == 1;

    private static bool IsC2paContinuationFragment(byte[] data, int payloadOffset, int payloadLength) =>
        payloadLength >= 8 && data[payloadOffset] == 0x4A && data[payloadOffset + 1] == 0x50 &&
        OfficeProvenanceBinary.ReadUInt32(data, payloadOffset + 4, littleEndian: false) > 1;

    private static void SortBySourceOffset<T>(List<T>? items) {
        if (items == null || items.Count < 2) return;
        IEnumerable<T> ordered = items.OrderBy(item => GetSourceOffset(item switch {
            OfficeProvenanceEvidence evidence => evidence.Location,
            OfficeProvenanceChange change => change.Location,
            _ => string.Empty
        }));
        T[] snapshot = ordered.ToArray();
        items.Clear();
        items.AddRange(snapshot);
    }

    private static int GetSourceOffset(string location) {
        int marker = location.IndexOf('@');
        if (marker < 0) return int.MaxValue;
        int value = 0;
        int index = marker + 1;
        bool found = false;
        while (index < location.Length && location[index] >= '0' && location[index] <= '9') {
            found = true;
            value = checked(value * 10 + location[index++] - '0');
        }
        return found ? value : int.MaxValue;
    }

    private static bool TryGetC2paSequence(
        byte[] data,
        int segmentStart,
        int payloadOffset,
        int payloadLength,
        OfficeProvenanceOptions options,
        ref int markerCount,
        out int sequenceEnd,
        out int manifestLength,
        out bool structurallyValid) {
        sequenceEnd = segmentStart;
        manifestLength = 0;
        structurallyValid = false;
        if (payloadLength < 8 || data[payloadOffset] != 0x4A || data[payloadOffset + 1] != 0x50 ||
            OfficeProvenanceBinary.ReadUInt32(data, payloadOffset + 4, littleEndian: false) != 1) return false;
        int firstFragmentLength = payloadLength - 8;
        bool completeFirstFragment = OfficeC2paManifestStore.IsValid(
            data, payloadOffset + 8, firstFragmentLength, options.MaxManifestBytes, options.MaxContainerEntries,
            out int declaredManifestLength);
        int boxHeaderLength = GetJumbfHeaderLength(data, payloadOffset + 8, firstFragmentLength);
        if (!completeFirstFragment &&
            (boxHeaderLength == 0 || !HasC2paDescriptionPrefix(data, payloadOffset + 8, firstFragmentLength, boxHeaderLength))) return false;
        bool hasDeclaredLength = completeFirstFragment || TryReadDeclaredJumbfLength(
            data,
            payloadOffset + 8,
            firstFragmentLength,
            options.MaxManifestBytes,
            out declaredManifestLength);

        long collected = firstFragmentLength;
        int current = payloadOffset + payloadLength;
        byte instanceHigh = data[payloadOffset + 2];
        byte instanceLow = data[payloadOffset + 3];
        uint expectedSequence = 2;
        while (!completeFirstFragment && (!hasDeclaredLength || collected < declaredManifestLength)) {
            if (!TryReadMarker(data, current, out byte marker, out int nextPayloadOffset, out int nextPayloadLength, out int nextEnd) ||
                marker != 0xEB || nextPayloadLength < 8 ||
                data[nextPayloadOffset] != 0x4A || data[nextPayloadOffset + 1] != 0x50 ||
                data[nextPayloadOffset + 2] != instanceHigh || data[nextPayloadOffset + 3] != instanceLow ||
                OfficeProvenanceBinary.ReadUInt32(data, nextPayloadOffset + 4, littleEndian: false) != expectedSequence) break;
            ReserveMarker(ref markerCount, options.MaxContainerEntries);
            int fragmentLength = nextPayloadLength - 8;
            collected += fragmentLength;
            current = nextEnd;
            expectedSequence++;
        }
        sequenceEnd = current;
        manifestLength = hasDeclaredLength ? declaredManifestLength : checked((int)Math.Min(collected, int.MaxValue));
        structurallyValid = completeFirstFragment;
        if (!structurallyValid && hasDeclaredLength && collected == declaredManifestLength) {
            byte[] reassembled = new byte[declaredManifestLength];
            Buffer.BlockCopy(data, payloadOffset + 8, reassembled, 0, firstFragmentLength);
            int destinationOffset = firstFragmentLength;
            int fragmentOffset = payloadOffset + payloadLength;
            while (fragmentOffset < sequenceEnd) {
                if (!TryReadMarker(data, fragmentOffset, out _, out int nextPayloadOffset, out int nextPayloadLength, out int nextEnd)) {
                    throw new InvalidDataException("JPEG APP11 sequence changed during bounded reassembly.");
                }
                int fragmentLength = nextPayloadLength - 8;
                Buffer.BlockCopy(data, nextPayloadOffset + 8, reassembled, destinationOffset, fragmentLength);
                destinationOffset += fragmentLength;
                fragmentOffset = nextEnd;
            }
            structurallyValid = OfficeC2paManifestStore.IsValid(
                reassembled, 0, reassembled.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
        }
        return true;
    }

    private static bool TryReadDeclaredJumbfLength(byte[] data, int offset, int available, long maximum, out int length) {
        length = 0;
        if (available < 8 || !OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "jumb")) return false;
        uint value = OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: false);
        if (value == 0) return false;
        if (value == 1) {
            if (available < 16) return false;
            ulong extended = OfficeProvenanceBinary.ReadUInt64(data, offset + 8, littleEndian: false);
            if (extended > (ulong)maximum || extended > int.MaxValue) return false;
            length = (int)extended;
            return length >= 46;
        }
        if (value > maximum || value > int.MaxValue) return false;
        length = (int)value;
        return length >= 38;
    }

    private static int GetJumbfHeaderLength(byte[] data, int offset, int available) {
        if (available < 8 || !OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "jumb")) return 0;
        return OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: false) == 1 && available >= 16 ? 16 : 8;
    }

    private static bool HasC2paDescriptionPrefix(byte[] data, int offset, int available, int boxHeaderLength) {
        int descriptionOffset = offset + boxHeaderLength;
        if (boxHeaderLength is not (8 or 16) || available < boxHeaderLength + 8 ||
            !OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "jumb") ||
            !OfficeProvenanceBinary.MatchesAscii(data, descriptionOffset + 4, "jumd")) return false;
        uint descriptionLength32 = OfficeProvenanceBinary.ReadUInt32(data, descriptionOffset, littleEndian: false);
        int descriptionHeaderLength;
        ulong descriptionLength;
        if (descriptionLength32 == 1) {
            if (available < boxHeaderLength + 16) return false;
            descriptionHeaderLength = 16;
            descriptionLength = OfficeProvenanceBinary.ReadUInt64(data, descriptionOffset + 8, littleEndian: false);
        } else {
            if (descriptionLength32 == 0) return false;
            descriptionHeaderLength = 8;
            descriptionLength = descriptionLength32;
        }
        if (descriptionLength < (ulong)(descriptionHeaderLength + 22) ||
            available < boxHeaderLength + descriptionHeaderLength + 22) return false;
        byte[] uuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        int descriptionPayloadOffset = descriptionOffset + descriptionHeaderLength;
        for (int index = 0; index < uuid.Length; index++) if (data[descriptionPayloadOffset + index] != uuid[index]) return false;
        int togglesOffset = descriptionPayloadOffset + uuid.Length;
        return (data[togglesOffset] & 0x02) != 0 &&
            OfficeProvenanceBinary.MatchesAscii(data, togglesOffset + 1, "c2pa") &&
            data[togglesOffset + 5] == 0;
    }

    internal static bool TryReadMarker(byte[] data, int start, out byte marker, out int payloadOffset, out int payloadLength, out int end) {
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

    private static int FindNextCompleteStart(byte[] data, int offset, int maximumEntries) {
        int lookaheadMarkers = 0;
        int candidate = FindNextStart(data, offset);
        while (candidate >= 0) {
            ReserveMarker(ref lookaheadMarkers, maximumEntries);
            if (TryFindCompleteImageEnd(data, candidate, ref lookaheadMarkers, maximumEntries, out _)) return candidate;
            candidate = FindNextStart(data, candidate + 2);
        }
        return -1;
    }

    private static bool TryFindCompleteImageEnd(
        byte[] data,
        int imageStart,
        ref int markerCount,
        int maximumEntries,
        out int imageEnd) {
        imageEnd = imageStart;
        int offset = imageStart + 2;
        while (offset < data.Length) {
            if (!TryReadMarker(data, offset, out byte marker, out _, out _, out int segmentEnd)) return false;
            ReserveMarker(ref markerCount, maximumEntries);
            if (marker == 0xD8) return false;
            if (marker == 0xD9) {
                imageEnd = segmentEnd;
                return true;
            }
            if (marker == 0xDA) return TryFindCompleteScanEnd(data, offset, ref markerCount, maximumEntries, out imageEnd);
            offset = segmentEnd;
        }
        return false;
    }

    private static bool TryFindCompleteScanEnd(
        byte[] data,
        int scanStart,
        ref int markerCount,
        int maximumEntries,
        out int imageEnd) {
        imageEnd = scanStart;
        if (!TryReadMarker(data, scanStart, out byte marker, out _, out _, out int offset) || marker != 0xDA) return false;
        while (offset < data.Length - 1) {
            if (data[offset] != 0xFF) { offset++; continue; }
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) return false;
            byte value = data[offset++];
            if (value == 0x00 || value == 0x01 || (value >= 0xD0 && value <= 0xD7)) continue;
            ReserveMarker(ref markerCount, maximumEntries);
            if (value == 0xD9) {
                imageEnd = offset;
                return true;
            }
            if (value == 0xD8 || data.Length - offset < 2) return false;
            int length = (data[offset] << 8) | data[offset + 1];
            if (length < 2 || length > data.Length - offset) return false;
            offset += length;
        }
        return false;
    }

    private static int FindEnd(byte[] data, int scanStart, ref int markerCount, int maximumEntries) {
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
            ReserveMarker(ref markerCount, maximumEntries);
            if (value == 0xD9) return offset;
            if (value == 0xD8) throw new InvalidDataException($"Unexpected JPEG SOI marker in scan data at offset {markerOffset}.");
            if (data.Length - offset < 2) break;
            int length = (data[offset] << 8) | data[offset + 1];
            if (length < 2 || length > data.Length - offset) break;
            offset += length;
        }
        throw new InvalidDataException("JPEG scan does not contain an end marker.");
    }

    private static void ReserveMarker(ref int markerCount, int maximumEntries) {
        if (markerCount >= maximumEntries) {
            throw new InvalidDataException($"The JPEG exceeds the configured container entry limit of {maximumEntries}.");
        }
        markerCount++;
    }
}
