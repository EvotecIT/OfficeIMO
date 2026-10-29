using OfficeIMO.Core.Internal;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>
/// Synthesizes native records without reading a source PUB. This is an experimental
/// hypothesis for native qualification, not an accepted Publisher writer profile.
/// </summary>
internal static class CandidateWriter {
    private const uint WidthEmu = 612 * 12700;
    private const uint HeightEmu = 792 * 12700;
    private const uint DocumentId = 0x100;
    internal static readonly Guid PublicationClassId = new("00021201-0000-0000-00c0-000000000046");
    private static readonly Guid QuillSubClassId = new("08c8f6da-969d-11d1-8e02-00c04fb6fece");

    internal static byte[] Create(string kind, string? value) {
        if (kind is not ("blank" or "text" or "shape" or "image")) throw new ArgumentException("Unknown candidate kind.");
        string? text = kind == "text" ? value ?? "Template-free Publisher writer probe\r" : null;
        if (text != null && !text.EndsWith('\r')) text += '\r';
        if (text?.Length > 4096) throw new ArgumentException("Probe text is limited to 4,096 characters.");
        byte[]? image = null;
        if (kind == "image") {
            if (value == null) throw new ArgumentException("The image candidate requires a PNG or JPEG path.");
            if (new FileInfo(value).Length > 1024 * 1024) throw new ArgumentException("Probe images are limited to 1 MiB.");
            image = File.ReadAllBytes(value);
        }
        NativePublicationGraph.Graph graph = NativePublicationGraph.Create(kind, DocumentId, WidthEmu, HeightEmu);
        var streams = new List<OfficeCompoundStream> {
            new OfficeCompoundStream("Contents", Contents(kind, graph)),
            new OfficeCompoundStream("Quill/QuillSub/CONTENTS", QuillCandidate.Create(text)),
            new OfficeCompoundStream("Escher/EscherStm", CandidateOfficeArt.Create(kind, image, graph.ShapeId, WidthEmu, HeightEmu)),
            new OfficeCompoundStream("Escher/EscherDelayStm", CandidateOfficeArt.CreateImageDelayStream(image)),
            new OfficeCompoundStream("\u0003Internal", InternalHeader()),
            new OfficeCompoundStream("\u0001CompObj", OfficeOleCompoundObjectWriter.Write(PublicationClassId,
                "Microsoft Publisher 3.0", "MSPublisher.3", string.Empty))
        };
        return WriteCompound(streams);
    }

    private static byte[] InternalHeader() => Bytes(writer => {
        writer.Write(0x00010000U); // Required version, checked after an exact 12-byte read.
        writer.Write((ushort)0); // Reserved +4: zero in native writer initialization.
        writer.Write(0U); // Context flags at unaligned +6: plain, inert publication.
        writer.Write((ushort)0); // Following-byte count +10: zero for the plain header.
    });

    private static byte[] WriteCompound(IReadOnlyList<OfficeCompoundStream> streams) {
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
        return OfficeCompoundFileWriter.Rewrite(generated, new Dictionary<string, byte[]>());
    }

    private static byte[] Contents(string kind, NativePublicationGraph.Graph graph) {
        IReadOnlyList<NativePublicationGraph.NativeRecord> chunks = graph.Chunks;
        byte[] header = NativeSchemaPrefix.Create(kind, DocumentId, 0, 0);
        var offsets = new List<int>();
        int offset = header.Length;
        foreach (var chunk in chunks) { offsets.Add(offset); offset += chunk.Bytes.Length; }
        int trailer = offset;
        var slots = Enumerable.Range(0, (int)DocumentId).Select(_ => Marker(0, 0x78)).ToList();
        slots.AddRange(chunks.Select((chunk, index) =>
            Variable(0, 0x88,
                ShortScalar(2, 0x18, checked((ushort)chunk.Kind)),
                Scalar(4, 0xB8, (uint)offsets[index]),
                chunk.Parent.HasValue ? Scalar(5, 0x68, chunk.Parent.Value) : Array.Empty<byte>())));
        byte[] directory = Variable(3, 0x90, slots.ToArray());
        header[0] = 0xE8; header[1] = 0xAC; header[2] = 0x2C;
        BitConverter.GetBytes(trailer).CopyTo(header, 0x1A);
        byte[] bytes = Join(header, Join(chunks.Select(chunk => chunk.Bytes).ToArray()),
            Chunk(Scalar(1, 0x20, DocumentId + (uint)chunks.Count), Scalar(2, 0x20, DocumentId + (uint)chunks.Count - 1), directory));
        byte[] prefix = NativeSchemaPrefix.Create(kind, DocumentId, bytes.Length, trailer);
        if (prefix.Length != header.Length) throw new InvalidDataException("Native schema fixup changed its declared prefix size.");
        prefix.CopyTo(bytes, 0);
        return bytes;
    }

}
