using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfOptimizationFileAssembler {
    private const int ObjectStreamChunkSize = 100;

    internal static byte[] Assemble(IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> objectStreamEligibility, int catalogId, int infoId, PdfFileVersion fileVersion, PdfOptimizationOptions options) {
        if (bodies.Count != objectStreamEligibility.Count) throw new ArgumentException("Object body and eligibility counts must match.", nameof(objectStreamEligibility));
        if (!options.UseObjectStreams && options.XrefFormat == PdfOptimizationXrefFormat.ClassicTable) {
            var objects = new List<byte[]>(bodies.Count);
            for (int i = 0; i < bodies.Count; i++) objects.Add(PdfObjectBytes.WrapIndirectObject(i + 1, bodies[i]));
            return PdfFileAssembler.Assemble(objects, catalogId, infoId, fileVersion);
        }
        return AssembleXrefStream(bodies, objectStreamEligibility, catalogId, infoId, fileVersion, options.UseObjectStreams);
    }

    private static byte[] AssembleXrefStream(IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> eligibility, int catalogId, int infoId, PdfFileVersion fileVersion, bool useObjectStreams) {
        fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, PdfFileVersion.Pdf15);
        var packs = BuildObjectStreamPacks(bodies, eligibility, useObjectStreams);
        int baseCount = bodies.Count;
        for (int i = 0; i < packs.Count; i++) packs[i].ObjectNumber = baseCount + i + 1;
        int xrefObjectNumber = baseCount + packs.Count + 1;
        int size = xrefObjectNumber + 1;
        var types = new byte[size]; var field2 = new long[size]; var field3 = new int[size];
        field3[0] = 65535;
        foreach (ObjectStreamPack pack in packs) for (int i = 0; i < pack.ObjectIds.Count; i++) { int id = pack.ObjectIds[i]; types[id] = 2; field2[id] = pack.ObjectNumber; field3[id] = i; }

        using var output = new MemoryStream();
        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-" + PdfFileAssembler.GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n"); output.Write(header, 0, header.Length);
        for (int id = 1; id <= baseCount; id++) {
            if (types[id] == 2) continue;
            types[id] = 1; field2[id] = output.Position;
            Write(output, PdfObjectBytes.WrapIndirectObject(id, bodies[id - 1]));
        }
        foreach (ObjectStreamPack pack in packs) {
            types[pack.ObjectNumber] = 1; field2[pack.ObjectNumber] = output.Position;
            byte[] content = BuildObjectStreamContent(pack, bodies, out int first); byte[] compressed = CompressFlate(content);
            string dictionary = "<< /Type /ObjStm /N " + pack.ObjectIds.Count.ToString(CultureInfo.InvariantCulture) + " /First " + first.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " >>";
            Write(output, PdfObjectBytes.WrapStreamObject(pack.ObjectNumber, dictionary, compressed));
        }
        long xrefOffset = output.Position; types[xrefObjectNumber] = 1; field2[xrefObjectNumber] = xrefOffset;
        byte[] xrefData = BuildXrefData(types, field2, field3);
        string xrefDictionary = "<< /Type /XRef /Size " + size.ToString(CultureInfo.InvariantCulture) + " /W [1 8 4] /Root " + PdfSyntaxEscaper.IndirectReference(catalogId) + (infoId > 0 ? " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) : string.Empty) + " /Length " + xrefData.Length.ToString(CultureInfo.InvariantCulture) + " >>";
        Write(output, PdfObjectBytes.WrapStreamObject(xrefObjectNumber, xrefDictionary, xrefData));
        Write(output, PdfEncoding.Latin1GetBytes("startxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n"));
        return output.ToArray();
    }

    private static List<ObjectStreamPack> BuildObjectStreamPacks(IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> eligibility, bool enabled) {
        var packs = new List<ObjectStreamPack>(); if (!enabled) return packs;
        ObjectStreamPack? current = null;
        for (int i = 0; i < bodies.Count; i++) {
            if (!eligibility[i]) continue;
            if (current is null || current.ObjectIds.Count == ObjectStreamChunkSize) { current = new ObjectStreamPack(); packs.Add(current); }
            current.ObjectIds.Add(i + 1);
        }
        return packs;
    }

    private static byte[] BuildObjectStreamContent(ObjectStreamPack pack, IReadOnlyList<byte[]> bodies, out int first) {
        var header = new StringBuilder(); int offset = 0;
        for (int i = 0; i < pack.ObjectIds.Count; i++) { int id = pack.ObjectIds[i]; header.Append(id.ToString(CultureInfo.InvariantCulture)).Append(' ').Append(offset.ToString(CultureInfo.InvariantCulture)).Append(' '); offset += bodies[id - 1].Length + 1; }
        header.Append('\n'); byte[] headerBytes = PdfEncoding.Latin1GetBytes(header.ToString()); first = headerBytes.Length;
        using var output = new MemoryStream(); output.Write(headerBytes, 0, headerBytes.Length);
        for (int i = 0; i < pack.ObjectIds.Count; i++) { byte[] body = bodies[pack.ObjectIds[i] - 1]; output.Write(body, 0, body.Length); output.WriteByte((byte)'\n'); }
        return output.ToArray();
    }

    private static byte[] BuildXrefData(byte[] types, long[] field2, int[] field3) {
        var result = new byte[types.Length * 13];
        for (int i = 0; i < types.Length; i++) { int offset = i * 13; result[offset] = types[i]; WriteBigEndian(result, offset + 1, field2[i], 8); WriteBigEndian(result, offset + 9, field3[i], 4); }
        return result;
    }

    private static void WriteBigEndian(byte[] destination, int offset, long value, int length) { for (int i = length - 1; i >= 0; i--) { destination[offset + i] = (byte)(value & 0xFF); value >>= 8; } }
    private static void Write(Stream output, byte[] bytes) => output.Write(bytes, 0, bytes.Length);
    private static byte[] CompressFlate(byte[] data) { using var output = new MemoryStream(); output.WriteByte(0x78); output.WriteByte(0x9C); using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true)) deflate.Write(data, 0, data.Length); uint adler = Adler32(data); output.WriteByte((byte)(adler >> 24)); output.WriteByte((byte)(adler >> 16)); output.WriteByte((byte)(adler >> 8)); output.WriteByte((byte)adler); return output.ToArray(); }
    private static uint Adler32(byte[] data) { const uint mod = 65521; uint a = 1, b = 0; for (int i = 0; i < data.Length; i++) { a = (a + data[i]) % mod; b = (b + a) % mod; } return (b << 16) | a; }

    private sealed class ObjectStreamPack { internal int ObjectNumber { get; set; } internal List<int> ObjectIds { get; } = new List<int>(); }
}
