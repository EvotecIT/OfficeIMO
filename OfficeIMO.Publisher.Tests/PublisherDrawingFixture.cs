using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Tests;

// Mutates declared records in a real publication. These inputs exercise the
// codec contract; they are not independently produced Publisher documents.
internal static class PublisherDrawingFixture {
    internal static byte[] Mutate(Dictionary<ushort, uint> values, int shapeType,
        Dictionary<ushort, byte[]>? complexValues = null, string fixture = "Simple.pub", uint objectId = 293,
        bool arrayLengthsExcludeHeader = false) =>
        Mutate(File.ReadAllBytes(PublisherNativeTests.Fixture(fixture)), values, shapeType, complexValues, objectId, arrayLengthsExcludeHeader);

    internal static byte[] Mutate(byte[] publication, Dictionary<ushort, uint> values, int shapeType,
        Dictionary<ushort, byte[]>? complexValues = null, uint objectId = 293, bool arrayLengthsExcludeHeader = false) {
        Assert.True(OfficeCompoundFileReader.TryRead(publication,
            out OfficeCompoundFile? source, out string? error), error);
        bool found = false;
        byte[] Rewrite(byte[] bytes, int start, int end) {
            using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
            for (int offset = start; offset < end;) {
                ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
                int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
                byte[] body = bytes.Skip(content).Take(boundary - content).ToArray();
                if (kind == 0xF004) {
                    for (int client = content; client < boundary;) {
                        int length = checked((int)BitConverter.ToUInt32(bytes, client + 4));
                        if (BitConverter.ToUInt16(bytes, client + 2) == 0xF011 && length == 10
                            && BitConverter.ToUInt32(bytes, client + 14) == objectId) {
                            found = true; body = RewriteShape(bytes, content, boundary, values, shapeType, complexValues, arrayLengthsExcludeHeader); break;
                        }
                        client += 8 + length;
                    }
                } else if ((initial & 15) == 15) body = Rewrite(bytes, content, boundary);
                writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
                offset = boundary;
                if (kind is 0xF000 or 0xF002 && boundary < end) { writer.Write(bytes, boundary, 4); offset += 4; }
            }
            return output.ToArray();
        }
        byte[] escher = source!.Streams["Escher/EscherStm"];
        byte[] replacement = Rewrite(escher, 0, escher.Length);
        Assert.True(found, "The requested fixture object must exist.");
        return OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = replacement });
    }

    private static byte[] RewriteShape(byte[] bytes, int start, int end, Dictionary<ushort, uint> values,
        int shapeType, Dictionary<ushort, byte[]>? complexValues, bool arrayLengthsExcludeHeader) {
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        bool wroteProperties = false;
        for (int offset = start; offset < end;) {
            ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
            int content = offset + 8, length = checked((int)BitConverter.ToUInt32(bytes, offset + 4));
            byte[] body = bytes.Skip(content).Take(length).ToArray();
            if (kind == 0xF00A) initial = (ushort)((shapeType << 4) | (initial & 15));
            if (kind is 0xF00B or 0xF122) {
                var entries = OfficeArtPropertyTableReader.Read(body, (ushort)(initial >> 4))
                    .Where(item => !values.ContainsKey(item.PropertyId) && complexValues?.ContainsKey(item.PropertyId) != true)
                    .Select(item => (Op: item.RawOperationId, Value: item.Value, Data: item.CopyComplexData())).ToList();
                entries.AddRange(values.Select(item => (item.Key, item.Value, (byte[]?)null)));
                if (complexValues != null)
                    entries.AddRange(complexValues.Select(item => ((ushort)(item.Key | 0x8000),
                        (uint)(item.Value.Length - (arrayLengthsExcludeHeader && item.Key is 0x145 or 0x146 or 0x197 ? 6 : 0)),
                        (byte[]?)item.Value)));
                entries = entries.OrderBy(item => item.Op & 0x3FFF).ToList();
                using var data = new MemoryStream(); using var properties = new BinaryWriter(data);
                foreach (var entry in entries) { properties.Write(entry.Op); properties.Write(entry.Value); }
                foreach (var entry in entries) if (entry.Data != null) properties.Write(entry.Data);
                body = data.ToArray(); initial = (ushort)((entries.Count << 4) | (initial & 15));
                wroteProperties = true;
            }
            writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
            offset = content + length;
        }
        Assert.True(wroteProperties);
        return output.ToArray();
    }
}
