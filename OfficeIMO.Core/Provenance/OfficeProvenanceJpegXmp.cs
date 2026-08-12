using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Provenance;

internal sealed class OfficeProvenanceJpegXmpResult {
    internal HashSet<int> SegmentStarts { get; } = new HashSet<int>();
    internal Dictionary<int, byte[]> Replacements { get; } = new Dictionary<int, byte[]>();
}

internal static class OfficeProvenanceJpegXmp {
    private static readonly byte[] StandardHeader = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
    private static readonly byte[] ExtendedHeader = Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
    private static readonly XNamespace XmpNoteNamespace = "http://ns.adobe.com/xmp/note/";
    private const int GuidLength = 32;
    private const int ExtendedChunkMetadataLength = GuidLength + 8;
    private const int MaximumSegmentPayload = ushort.MaxValue - 2;

    internal static OfficeProvenanceJpegXmpResult ProcessImage(
        byte[] data,
        int start,
        int imageIndex,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes) {
        var result = new OfficeProvenanceJpegXmpResult();
        var standards = new List<StandardPacket>();
        var extensions = new List<ExtendedChunk>();
        int offset = start;
        while (offset < data.Length && OfficeProvenanceJpeg.TryReadMarker(
            data, offset, out byte marker, out int payloadOffset, out int payloadLength, out int segmentEnd)) {
            if (marker is 0xDA or 0xD9) break;
            if (marker == 0xE1 && Matches(data, payloadOffset, payloadLength, StandardHeader)) {
                int packetLength = payloadLength - StandardHeader.Length;
                byte[] packet = new byte[packetLength];
                Buffer.BlockCopy(data, payloadOffset + StandardHeader.Length, packet, 0, packetLength);
                standards.Add(new StandardPacket(offset, packet));
                result.SegmentStarts.Add(offset);
            } else if (marker == 0xE1 && TryReadExtendedChunk(data, offset, payloadOffset, payloadLength, out ExtendedChunk? chunk)) {
                extensions.Add(chunk!);
                result.SegmentStarts.Add(offset);
            }
            offset = segmentEnd;
        }

        var references = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        var currentPackets = standards.ToDictionary(item => item.SegmentStart, item => item.Packet);
        foreach (StandardPacket standard in standards) {
            string location = $"JPEG[{imageIndex}]/APP1-XMP@{standard.SegmentStart}";
            if (context != null) OfficeProvenanceXmp.Inspect(standard.Packet, options, context, location);
            foreach (string guid in GetExtendedGuids(standard.Packet, options)) {
                if (!references.TryGetValue(guid, out List<int>? starts)) references.Add(guid, starts = new List<int>());
                starts.Add(standard.SegmentStart);
            }
            if (removalOptions != null && changes != null &&
                OfficeProvenanceXmp.TryRemoveAiDeclarations(standard.Packet, removalOptions, location, changes, out byte[] cleaned)) {
                currentPackets[standard.SegmentStart] = cleaned;
            }
        }

        foreach (KeyValuePair<string, List<int>> reference in references) {
            ExtendedChunk[] chunks = extensions
                .Where(item => string.Equals(item.Guid, reference.Key, StringComparison.OrdinalIgnoreCase))
                .OrderBy(item => item.Offset)
                .ToArray();
            string location = $"JPEG[{imageIndex}]/APP1-ExtendedXMP[{reference.Key}]";
            if (!TryAssemble(chunks, options.MaxAssetBytes, out byte[] packet)) {
                if (chunks.Length != 0) context?.Diagnostics.Add($"{location}: extended XMP chunks are incomplete or malformed.");
                continue;
            }
            if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
            if (removalOptions == null || changes == null) continue;

            var pendingChanges = new List<OfficeProvenanceChange>();
            if (!OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, pendingChanges, out byte[] cleanedPacket)) continue;
            string replacementGuid = ComputeGuid(cleanedPacket);
            var updatedStandards = new Dictionary<int, byte[]>();
            bool referencesUpdated = true;
            foreach (int standardStart in reference.Value.Distinct()) {
                if (!TryReplaceExtendedGuid(currentPackets[standardStart], reference.Key, replacementGuid, options, out byte[] updated)) {
                    referencesUpdated = false;
                    break;
                }
                updatedStandards[standardStart] = updated;
            }
            if (!referencesUpdated) continue;

            foreach (KeyValuePair<int, byte[]> updated in updatedStandards) currentPackets[updated.Key] = updated.Value;
            ApplyExtendedReplacement(result.Replacements, chunks, replacementGuid, cleanedPacket);
            changes.AddRange(pendingChanges);
        }

        foreach (StandardPacket standard in standards) {
            byte[] current = currentPackets[standard.SegmentStart];
            if (!ReferenceEquals(current, standard.Packet)) {
                result.Replacements[standard.SegmentStart] = CreateSegment(Join(StandardHeader, current));
            }
        }
        return result;
    }

    private static bool TryReadExtendedChunk(
        byte[] data,
        int segmentStart,
        int payloadOffset,
        int payloadLength,
        out ExtendedChunk? chunk) {
        chunk = null;
        if (!Matches(data, payloadOffset, payloadLength, ExtendedHeader) ||
            payloadLength < ExtendedHeader.Length + ExtendedChunkMetadataLength) return false;
        int metadataOffset = payloadOffset + ExtendedHeader.Length;
        string guid = Encoding.ASCII.GetString(data, metadataOffset, GuidLength);
        if (guid.Length != GuidLength || guid.Any(character => !Uri.IsHexDigit(character))) return false;
        uint fullLength = OfficeProvenanceBinary.ReadUInt32(data, metadataOffset + GuidLength, littleEndian: false);
        uint chunkOffset = OfficeProvenanceBinary.ReadUInt32(data, metadataOffset + GuidLength + 4, littleEndian: false);
        int dataOffset = metadataOffset + ExtendedChunkMetadataLength;
        int dataLength = payloadOffset + payloadLength - dataOffset;
        byte[] bytes = new byte[dataLength];
        Buffer.BlockCopy(data, dataOffset, bytes, 0, dataLength);
        chunk = new ExtendedChunk(segmentStart, guid, fullLength, chunkOffset, bytes);
        return true;
    }

    private static bool TryAssemble(ExtendedChunk[] chunks, long maximumBytes, out byte[] packet) {
        packet = Array.Empty<byte>();
        if (chunks.Length == 0) return false;
        uint fullLength = chunks[0].FullLength;
        if (fullLength == 0 || fullLength > maximumBytes || fullLength > int.MaxValue ||
            chunks.Any(item => item.FullLength != fullLength)) return false;
        long cursor = 0;
        foreach (ExtendedChunk chunk in chunks) {
            if (chunk.Offset != cursor || chunk.Data.LongLength > fullLength - cursor) return false;
            cursor += chunk.Data.LongLength;
        }
        if (cursor != fullLength) return false;
        packet = new byte[(int)fullLength];
        cursor = 0;
        foreach (ExtendedChunk chunk in chunks) {
            Buffer.BlockCopy(chunk.Data, 0, packet, (int)cursor, chunk.Data.Length);
            cursor += chunk.Data.LongLength;
        }
        return true;
    }

    private static IEnumerable<string> GetExtendedGuids(byte[] packet, OfficeProvenanceOptions options) {
        if (!TryLoad(packet, options, out XDocument? document) || document == null) yield break;
        foreach (XObject node in FindGuidNodes(document)) {
            string value = node is XAttribute attribute ? attribute.Value : ((XElement)node).Value;
            string guid = value.Trim();
            if (guid.Length == GuidLength && guid.All(Uri.IsHexDigit)) yield return guid;
        }
    }

    private static bool TryReplaceExtendedGuid(
        byte[] packet,
        string oldGuid,
        string newGuid,
        OfficeProvenanceOptions options,
        out byte[] updated) {
        updated = packet;
        if (!TryLoad(packet, options, out XDocument? document) || document == null) return false;
        bool changed = false;
        foreach (XObject node in FindGuidNodes(document)) {
            string value = node is XAttribute attribute ? attribute.Value : ((XElement)node).Value;
            if (!string.Equals(value.Trim(), oldGuid, StringComparison.OrdinalIgnoreCase)) continue;
            if (node is XAttribute matchingAttribute) matchingAttribute.Value = newGuid;
            else ((XElement)node).Value = newGuid;
            changed = true;
        }
        if (!changed) return false;
        updated = Serialize(document);
        return true;
    }

    private static IEnumerable<XObject> FindGuidNodes(XDocument document) {
        foreach (XElement element in document.Descendants()) {
            if (element.Name == XmpNoteNamespace + "HasExtendedXMP") yield return element;
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.Name == XmpNoteNamespace + "HasExtendedXMP") yield return attribute;
            }
        }
    }

    private static bool TryLoad(byte[] packet, OfficeProvenanceOptions options, out XDocument? document) {
        document = null;
        if (packet.LongLength > options.MaxAssetBytes) return false;
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = options.MaxAssetBytes,
                MaxCharactersFromEntities = 0,
                IgnoreWhitespace = false
            };
            using var stream = new MemoryStream(packet, writable: false);
            using XmlReader reader = XmlReader.Create(stream, settings);
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            return document.Root != null;
        } catch (XmlException) {
            return false;
        }
    }

    private static byte[] Serialize(XDocument document) {
        using var stream = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(stream, settings)) document.Save(writer);
        return stream.ToArray();
    }

    private static void ApplyExtendedReplacement(
        Dictionary<int, byte[]> replacements,
        ExtendedChunk[] chunks,
        string guid,
        byte[] packet) {
        int maximumChunkBytes = MaximumSegmentPayload - ExtendedHeader.Length - ExtendedChunkMetadataLength;
        using var segments = new MemoryStream();
        for (int offset = 0; offset < packet.Length; offset += maximumChunkBytes) {
            int count = Math.Min(maximumChunkBytes, packet.Length - offset);
            byte[] metadata = new byte[ExtendedHeader.Length + ExtendedChunkMetadataLength + count];
            Buffer.BlockCopy(ExtendedHeader, 0, metadata, 0, ExtendedHeader.Length);
            Encoding.ASCII.GetBytes(guid, 0, guid.Length, metadata, ExtendedHeader.Length);
            WriteUInt32(metadata, ExtendedHeader.Length + GuidLength, (uint)packet.Length);
            WriteUInt32(metadata, ExtendedHeader.Length + GuidLength + 4, (uint)offset);
            Buffer.BlockCopy(packet, offset, metadata, ExtendedHeader.Length + ExtendedChunkMetadataLength, count);
            byte[] segment = CreateSegment(metadata);
            segments.Write(segment, 0, segment.Length);
        }
        replacements[chunks[0].SegmentStart] = segments.ToArray();
        for (int index = 1; index < chunks.Length; index++) replacements[chunks[index].SegmentStart] = Array.Empty<byte>();
    }

    private static string ComputeGuid(byte[] packet) {
        using MD5 algorithm = MD5.Create();
        byte[] digest = algorithm.ComputeHash(packet);
        var builder = new StringBuilder(GuidLength);
        foreach (byte value in digest) builder.Append(value.ToString("X2"));
        return builder.ToString();
    }

    private static byte[] CreateSegment(byte[] payload) {
        if (payload.Length > MaximumSegmentPayload) throw new InvalidDataException("JPEG XMP packet exceeds the APP1 segment limit.");
        byte[] segment = new byte[payload.Length + 4];
        segment[0] = 0xFF;
        segment[1] = 0xE1;
        int length = payload.Length + 2;
        segment[2] = (byte)(length >> 8);
        segment[3] = (byte)length;
        Buffer.BlockCopy(payload, 0, segment, 4, payload.Length);
        return segment;
    }

    private static byte[] Join(byte[] first, byte[] second) {
        byte[] result = new byte[first.Length + second.Length];
        Buffer.BlockCopy(first, 0, result, 0, first.Length);
        Buffer.BlockCopy(second, 0, result, first.Length, second.Length);
        return result;
    }

    private static bool Matches(byte[] data, int offset, int available, byte[] expected) {
        if (available < expected.Length || offset < 0 || expected.Length > data.Length - offset) return false;
        for (int index = 0; index < expected.Length; index++) if (data[offset + index] != expected[index]) return false;
        return true;
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private sealed class StandardPacket {
        internal StandardPacket(int segmentStart, byte[] packet) {
            SegmentStart = segmentStart;
            Packet = packet;
        }
        internal int SegmentStart { get; }
        internal byte[] Packet { get; }
    }

    private sealed class ExtendedChunk {
        internal ExtendedChunk(int segmentStart, string guid, uint fullLength, uint offset, byte[] data) {
            SegmentStart = segmentStart;
            Guid = guid;
            FullLength = fullLength;
            Offset = offset;
            Data = data;
        }
        internal int SegmentStart { get; }
        internal string Guid { get; }
        internal uint FullLength { get; }
        internal uint Offset { get; }
        internal byte[] Data { get; }
    }
}
