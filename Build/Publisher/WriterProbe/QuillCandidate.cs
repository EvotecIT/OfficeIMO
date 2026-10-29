using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Minimal Quill record hypothesis. Native required defaults remain subject to qualification.</summary>
internal static class QuillCandidate {
    private const int FirstPayloadOffset = 512;

    internal static byte[] Create(string? text) {
        if (text == null) return Join(Header(0, FirstPayloadOffset), new byte[FirstPayloadOffset - 32]);
        byte[] font = Bytes(writer => {
            writer.Write(0U); writer.Write(1U); writer.Write(new byte[12]); writer.Write(0U);
            writer.Write((ushort)5); writer.Write(Encoding.Unicode.GetBytes("Arial")); writer.Write(0U);
        });
        var records = new List<(string Tag, byte[] Body)> {
            ("TEXT", Encoding.Unicode.GetBytes(text)),
            ("SYID", Bytes(writer => { writer.Write(0U); writer.Write(1U); writer.Write(1U); })),
            ("STRS", Bytes(writer => { writer.Write(1U); writer.Write(4U); writer.Write((uint)text.Length); })),
            ("FONT", font)
        };
        // Formatting endpoints are absolute byte positions in the complete stream.
        int textEnd = FirstPayloadOffset + text.Length * 2;
        records.Add(("FDPC", Range(textEnd, Chunk(Scalar(0x0C, 0x20, 18 * 12700)))));
        records.Add(("FDPP", Range(textEnd, Chunk())));
        int logicalLength = FirstPayloadOffset + records.Sum(record => record.Body.Length);
        byte[] header = Header(records.Count, logicalLength);
        int directoryEnd = header.Length + records.Count * 24;
        if (directoryEnd > FirstPayloadOffset) throw new InvalidDataException("The Quill directory exceeds its allocated envelope.");
        int offset = FirstPayloadOffset;
        return Bytes(writer => {
            writer.Write(header);
            foreach (var record in records) {
                writer.Write((ushort)24); writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write((ushort)0); writer.Write((ushort)1); writer.Write((ushort)0);
                writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write(offset); writer.Write(record.Body.Length);
                offset += record.Body.Length;
            }
            writer.Write(new byte[FirstPayloadOffset - directoryEnd]);
            foreach (var record in records) writer.Write(record.Body);
        });
    }

    private static byte[] Header(int count, int logicalLength) => Bytes(writer => {
        writer.Write(Encoding.ASCII.GetBytes("CHNKINK "));
        // Native checks require version 4/7, byte +0x0F == 3 and a bounded,
        // seekable DWORD at +0x14 of at least 512. The remaining values follow
        // independently authored controls; their sufficiency is unqualified.
        writer.Write((ushort)4); writer.Write((ushort)7); writer.Write((ushort)count);
        writer.Write((byte)0); writer.Write((byte)3);
        writer.Write(FirstPayloadOffset); writer.Write(logicalLength);
        writer.Write((ushort)(FirstPayloadOffset - 8)); writer.Write((ushort)count); writer.Write(uint.MaxValue);
    });

    private static byte[] Range(int textEnd, byte[] style) => Bytes(writer => {
        writer.Write((ushort)1); writer.Write(new byte[6]);
        writer.Write(textEnd); writer.Write((ushort)14); writer.Write(style);
    });
}
