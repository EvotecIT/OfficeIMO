using OfficeIMO.Core.Internal;
using static OfficeIMO.Publisher.Internal.PublisherNativeCreationEncoding;

namespace OfficeIMO.Publisher.Internal;

/// <summary>
/// Synthesizes native records without reading a source PUB. This is an experimental
/// hypothesis for native qualification, not an accepted Publisher writer profile.
/// </summary>
internal static partial class PublisherNativeCreationCodec {
    private const uint DocumentId = 0x100;
    internal static readonly Guid PublicationClassId = new("00021201-0000-0000-00c0-000000000046");
    private static readonly Guid QuillSubClassId = new("08c8f6da-969d-11d1-8e02-00c04fb6fece");

    /// <summary>Creates candidate bytes under an independent operation budget; this does not admit public source-free saving.</summary>
    internal static byte[] Create(PublisherNativeCreationPlan plan, PublisherReadOptions options, long maxOutputBytes, CancellationToken token) {
        if (plan == null) throw new ArgumentNullException(nameof(plan));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (maxOutputBytes < 1 || maxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        token.ThrowIfCancellationRequested();
        var context = new PublisherParseContext(options.Clone(), token);
        ValidateResources(plan, context, maxOutputBytes);
        PublisherNativeCreationGraph.Graph graph = PublisherNativeCreationGraph.Create(plan, DocumentId, context);
        IReadOnlyList<PublisherNativeCreationArtwork.ImageBinding> images = PublisherNativeCreationArtwork.BindImages(graph, context);
        var streams = new List<OfficeCompoundStream> {
            new OfficeCompoundStream("Contents", Contents(graph, context)),
            new OfficeCompoundStream("Quill/QuillSub/CONTENTS", PublisherNativeCreationQuill.Create(graph, context)),
            new OfficeCompoundStream("Escher/EscherStm", PublisherNativeCreationArtwork.Create(graph, images, context)),
            new OfficeCompoundStream("Escher/EscherDelayStm", PublisherNativeCreationArtwork.CreateImageDelayStream(images, context)),
            new OfficeCompoundStream("\u0003Internal", InternalHeader()),
            new OfficeCompoundStream("\u0001CompObj", OfficeOleCompoundObjectWriter.Write(PublicationClassId,
                "Microsoft Publisher 3.0", "MSPublisher.3", string.Empty))
        };
        return WriteCompound(streams, maxOutputBytes, token);
    }

    private static byte[] InternalHeader() => Bytes(writer => {
        writer.Write(0x00010000U); // Required version, checked after an exact 12-byte read.
        writer.Write((ushort)0); // Reserved +4: zero in native writer initialization.
        writer.Write(0U); // Context flags at unaligned +6: plain, inert publication.
        writer.Write((ushort)0); // Following-byte count +10: zero for the plain header.
    });

    private static byte[] WriteCompound(IReadOnlyList<OfficeCompoundStream> streams, long maxOutputBytes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        // Compose a fresh directory model so the canonical writer carries the
        // observed nested storage identity. No publication is read as a seed.
        var root = new OfficeCompoundFileEntry("Root Entry", string.Empty, 5, 0, classId: PublicationClassId);
        var entries = new List<OfficeCompoundFileEntry> {
            root,
            // The preliminary native version reader opens this storage before
            // validation, including when the publication has no embedded objects.
            new("Objects", "Objects", 1, 0),
            new("Quill", "Quill", 1, 0),
            new("QuillSub", "Quill/QuillSub", 1, 0, classId: QuillSubClassId),
            new("Escher", "Escher", 1, 0)
        };
        entries.AddRange(streams.Select(stream => new OfficeCompoundFileEntry(
            stream.Name.Split('/').Last(), stream.Name, 2, stream.Length)));
        var generated = new OfficeCompoundFile(
            streams.ToDictionary(stream => stream.Name, stream => stream.Bytes, StringComparer.OrdinalIgnoreCase), entries, root);
        return OfficeCompoundFileWriter.Rewrite(generated, new Dictionary<string, byte[]>(),
            maxOutputBytes: maxOutputBytes, cancellationToken: token);
    }

    private static byte[] Contents(PublisherNativeCreationGraph.Graph graph, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        IReadOnlyList<PublisherNativeCreationGraph.NativeRecord> chunks = graph.Chunks;
        uint[] objectKinds = graph.Objects.Select(item => item.Source.NativeKind).Distinct().ToArray();
        byte[] header = PublisherNativeCreationSchema.Create(objectKinds, DocumentId, 0, 0);
        var offsets = new List<int>();
        int offset = header.Length;
        foreach (var chunk in chunks) {
            context.Token.ThrowIfCancellationRequested();
            offsets.Add(offset); offset = checked(offset + chunk.Bytes.Length);
        }
        int trailer = offset;
        var slots = Enumerable.Range(0, (int)DocumentId).Select(_ => new byte[] { 0, 0x78 }).ToList();
        for (int index = 0; index < chunks.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            var chunk = chunks[index];
            slots.Add(PublisherOplWriter.Nested(0, 0x88,
                PublisherOplWriter.UInt16(2, checked((ushort)chunk.Kind), 0x18),
                PublisherOplWriter.UInt32(4, (uint)offsets[index], 0xB8),
                chunk.Parent.HasValue ? PublisherOplWriter.UInt32(5, chunk.Parent.Value, 0x68) : Array.Empty<byte>()));
        }
        byte[] directory = PublisherOplWriter.Nested(3, 0x90, slots.ToArray());
        header[0] = 0xE8; header[1] = 0xAC; header[2] = 0x2C;
        BitConverter.GetBytes(trailer).CopyTo(header, 0x1A);
        context.Token.ThrowIfCancellationRequested();
        byte[] bytes = Join(header, Join(chunks.Select(chunk => chunk.Bytes).ToArray()),
            PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(1, DocumentId + (uint)chunks.Count, 0x20), PublisherOplWriter.UInt32(2, DocumentId + (uint)chunks.Count - 1, 0x20), directory));
        byte[] prefix = PublisherNativeCreationSchema.Create(objectKinds, DocumentId, bytes.Length, trailer);
        if (prefix.Length != header.Length) throw new InvalidDataException("Native schema fixup changed its declared prefix size.");
        prefix.CopyTo(bytes, 0);
        context.Token.ThrowIfCancellationRequested();
        return bytes;
    }

}
