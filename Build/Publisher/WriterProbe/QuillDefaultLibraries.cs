using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Constructs small Quill library records from recovered loader contracts.</summary>
internal static class QuillDefaultLibraries {
    internal static IEnumerable<QuillRecord> Create() {
        byte[] defaultName = Bytes(writer => {
            writer.Write((ushort)0); // Empty UTF-16 name.
            writer.Write(uint.MaxValue); // Select the native built-in default style.
        });
        byte[] emptyStyle = Bytes(writer => {
            writer.Write((ushort)2); // Two UTF-16 words carry the four-byte empty OPL.
            writer.Write(Chunk());
        });
        yield return new QuillRecord("STSH", 0, Indexed(4, defaultName));
        yield return new QuillRecord("STSH", 1, Indexed(0, emptyStyle, emptyStyle));
        yield return new QuillRecord("INK ", 0, Bytes(writer => {
            writer.Write((ushort)0x74); writer.Write((ushort)0x74);
        }));
        // Native default character mapping uses font index zero. Generate that
        // entry explicitly, without querying or copying machine font metadata.
        byte[] font = Bytes(writer => {
            writer.Write((ushort)5); writer.Write(System.Text.Encoding.Unicode.GetBytes("Arial"));
            writer.Write(0U);
        });
        yield return new QuillRecord("FONT", 0, Indexed(4, font));
    }

    private static byte[] Indexed(uint fixedTailBytes, params byte[][] entries) {
        int offset = entries.Length * 4;
        int bodyBytes = offset + entries.Sum(entry => entry.Length);
        return Bytes(writer => {
            writer.Write(bodyBytes); writer.Write(entries.Length);
            writer.Write(0U); // This header DWORD is ignored by the recovered reader.
            writer.Write(fixedTailBytes); // Low14bits specify the fixed per-entry tail.
            writer.Write(0U); // Optional runtime flag bit0 remains clear.
            foreach (byte[] entry in entries) {
                writer.Write(offset); offset += entry.Length;
            }
            foreach (byte[] entry in entries) writer.Write(entry);
        });
    }
}
