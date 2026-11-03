using static OfficeIMO.Publisher.Internal.PublisherNativeCreationEncoding;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Constructs small Quill library records from recovered loader contracts.</summary>
internal static class PublisherNativeCreationQuillDefaults {
    internal static IEnumerable<PublisherNativeCreationQuillRecord> Create(PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        byte[] defaultName = Bytes(writer => {
            writer.Write((ushort)0); // Empty UTF-16 name.
            writer.Write(uint.MaxValue); // Select the native built-in default style.
        });
        byte[] emptyStyle = Bytes(writer => {
            writer.Write((ushort)2); // Two UTF-16 words carry the four-byte empty OPL.
            writer.Write(PublisherOplWriter.Chunk());
        });
        yield return new PublisherNativeCreationQuillRecord("STSH", 0, Indexed(4, context, defaultName));
        yield return new PublisherNativeCreationQuillRecord("STSH", 1, Indexed(0, context, emptyStyle, emptyStyle));
        context.Token.ThrowIfCancellationRequested();
        yield return new PublisherNativeCreationQuillRecord("INK ", 0, Bytes(writer => {
            writer.Write((ushort)0x74); writer.Write((ushort)0x74);
        }));
        // Native default character mapping uses font index zero. Generate that
        // entry explicitly, without querying or copying machine font metadata.
        context.Token.ThrowIfCancellationRequested();
        byte[] font = Bytes(writer => {
            writer.Write((ushort)5); writer.Write(System.Text.Encoding.Unicode.GetBytes("Arial"));
            writer.Write(0U);
        });
        yield return new PublisherNativeCreationQuillRecord("FONT", 0, Indexed(4, context, font));
        context.Token.ThrowIfCancellationRequested();
    }

    private static byte[] Indexed(uint fixedTailBytes, PublisherParseContext context, params byte[][] entries) {
        context.Token.ThrowIfCancellationRequested();
        int offset = entries.Length * 4;
        int bodyLength = 0;
        foreach (byte[] entry in entries) {
            context.Token.ThrowIfCancellationRequested();
            bodyLength = checked(bodyLength + entry.Length);
        }
        int bodyBytes = offset + bodyLength;
        byte[] result = Bytes(writer => {
            writer.Write(bodyBytes); writer.Write(entries.Length);
            writer.Write(0U); // This header DWORD is ignored by the recovered reader.
            writer.Write(fixedTailBytes); // Low14bits specify the fixed per-entry tail.
            writer.Write(0U); // Optional runtime flag bit0 remains clear.
            foreach (byte[] entry in entries) {
                context.Token.ThrowIfCancellationRequested();
                writer.Write(offset); offset += entry.Length;
            }
            foreach (byte[] entry in entries) {
                context.Token.ThrowIfCancellationRequested();
                writer.Write(entry);
            }
        });
        context.Token.ThrowIfCancellationRequested();
        return result;
    }
}
