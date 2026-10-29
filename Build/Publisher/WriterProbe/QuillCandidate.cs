using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Minimal Quill record hypothesis. Native required defaults remain subject to qualification.</summary>
internal static class QuillCandidate {
    private const int FirstPayloadOffset = 512;
    private const int MinimumStreamBytes = 1024;

    internal static byte[] Create(string? text) {
        var records = new List<QuillRecord>();
        if (text != null) {
            records.Add(new QuillRecord("TEXT", 0, Encoding.Unicode.GetBytes(text)));
            records.Add(new QuillRecord("SYID", 0, Bytes(writer => { writer.Write(0U); writer.Write(1U); writer.Write(1U); })));
            records.Add(new QuillRecord("STRS", 0, Bytes(writer => { writer.Write(1U); writer.Write(4U); writer.Write((uint)text.Length); })));
        }
        records.AddRange(QuillDefaultLibraries.Create());
        if (text != null) {
            // TEXT is first, so its absolute formatting endpoint follows its
            // actual serialized position even when the other libraries grow.
            int textEnd = FirstPayloadOffset + text.Length * 2;
            records.Add(new QuillRecord("FDPC", 0, Range(textEnd, Chunk(Scalar(0x0C, 0x20, 18 * 12700)))));
            records.Add(new QuillRecord("FDPP", 0, Range(textEnd, Chunk())));
        }
        int payloadEnd = FirstPayloadOffset + records.Sum(record => record.Body.Length);
        int logicalLength = Math.Max(MinimumStreamBytes, payloadEnd);
        byte[] header = Header(records.Count, logicalLength);
        int directoryEnd = header.Length + records.Count * 24;
        if (directoryEnd > FirstPayloadOffset) throw new InvalidDataException("The Quill directory exceeds its allocated envelope.");
        int offset = FirstPayloadOffset;
        return Bytes(writer => {
            writer.Write(header);
            foreach (var record in records) {
                writer.Write((ushort)24); writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write(record.Id); writer.Write((ushort)1); writer.Write((ushort)0);
                writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write(offset); writer.Write(record.Body.Length);
                offset += record.Body.Length;
            }
            writer.Write(new byte[FirstPayloadOffset - directoryEnd]);
            foreach (var record in records) writer.Write(record.Body);
            writer.Write(new byte[logicalLength - payloadEnd]);
        });
    }

    private static byte[] Header(int totalDescriptors, int logicalLength) => Bytes(writer => {
        writer.Write(Encoding.ASCII.GetBytes("CHNKINK "));
        // Native checks require version 4/7, byte +0x0F == 3 and a bounded,
        // seekable DWORD at +0x14 of at least 512. The remaining values follow
        // independently authored controls; their sufficiency is unqualified.
        writer.Write((ushort)4); writer.Write((ushort)7); writer.Write((ushort)totalDescriptors);
        writer.Write((byte)0); writer.Write((byte)3);
        writer.Write(FirstPayloadOffset); writer.Write(logicalLength);
        // The probe emits one directory; its count equals the total in +0x0C.
        writer.Write((ushort)(FirstPayloadOffset - 8)); writer.Write((ushort)totalDescriptors); writer.Write(uint.MaxValue);
    });

    private static byte[] Range(int textEnd, byte[] style) => Bytes(writer => {
        writer.Write((ushort)1); writer.Write(new byte[6]);
        writer.Write(textEnd); writer.Write((ushort)14); writer.Write(style);
    });
}

internal sealed record QuillRecord(string Tag, ushort Id, byte[] Body);
