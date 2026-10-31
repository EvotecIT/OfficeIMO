using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Generates layout slots for native frame ordinals; frame geometry is applied by the native loader.</summary>
internal static class QuillLayoutCatalog {
    internal static byte[] Create(NativePublicationGraph.Graph graph) => Bytes(writer => {
        writer.Write(graph.Stories.Max(story => story.LayoutId));
        writer.Write((uint)graph.Stories.Count);
        foreach (var story in graph.Stories) writer.Write(story.LayoutId);
        foreach (var story in graph.Stories) {
            // Observed kind 0D metadata fields 200/201. Their encoding is
            // recovered; native hyphenation also depends on other story state.
            writer.Write(Chunk(Bytes(fields => {
                fields.Write((ushort)0x0A00);
                fields.Write((ushort)0x2201); fields.Write(18U * 12700);
            })));
            int count = story.Source.CellEnds?.Count ?? graph.Objects.Count(frame => frame.Story == story);
            writer.Write((uint)count);
            for (int index = 0; index < count; index++)
                writer.Write(Chunk()); // Kind 9 layout slot; no copied native cache or opaque bytes.
        }
    });
}
