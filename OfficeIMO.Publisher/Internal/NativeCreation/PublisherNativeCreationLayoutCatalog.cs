using static OfficeIMO.Publisher.Internal.PublisherNativeCreationEncoding;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Generates layout slots for native frame ordinals; frame geometry is applied by the native loader.</summary>
internal static class PublisherNativeCreationLayoutCatalog {
    internal static byte[] Create(PublisherNativeCreationGraph.Graph graph, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        uint maximum = graph.Stories.Max(story => {
            context.Token.ThrowIfCancellationRequested();
            return story.LayoutId;
        });
        byte[] result = Bytes(writer => {
            writer.Write(maximum);
            writer.Write((uint)graph.Stories.Count);
            foreach (var story in graph.Stories) {
                context.Token.ThrowIfCancellationRequested();
                writer.Write(story.LayoutId);
            }
            foreach (var story in graph.Stories) {
                context.Token.ThrowIfCancellationRequested();
                // Observed kind 0D metadata fields 200/201. Their encoding is
                // recovered; native hyphenation also depends on other story state.
                writer.Write(PublisherOplWriter.Chunk(Bytes(fields => {
                    fields.Write((ushort)0x0A00);
                    fields.Write((ushort)0x2201); fields.Write(18U * 12700);
                })));
                int count = story.Source.CellEnds?.Count ?? graph.Objects.Count(frame => {
                    context.Token.ThrowIfCancellationRequested();
                    return frame.Story == story;
                });
                writer.Write((uint)count);
                for (int index = 0; index < count; index++) {
                    context.Token.ThrowIfCancellationRequested();
                    writer.Write(PublisherOplWriter.Chunk()); // Kind 9 layout slot; no copied native cache or opaque bytes.
                }
            }
        });
        context.Token.ThrowIfCancellationRequested();
        return result;
    }
}
