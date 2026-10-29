using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>
/// Synthesizes native records without reading a source PUB. This is an experimental
/// hypothesis for native qualification, not an accepted Publisher writer profile.
/// </summary>
internal static class CandidateWriter {
    private const uint WidthEmu = 612 * 12700;
    private const uint HeightEmu = 792 * 12700;

    internal static byte[] Create(string kind, string? value, Guid? rootClassId = null, bool nativeHeader = false, bool bootstrapMetadata = false, bool nativeSchema = false, bool pageRoles = false) {
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
        uint baseId = bootstrapMetadata ? 0x100U : 0U;
        var streams = new List<OfficeCompoundStream> {
            new OfficeCompoundStream("Contents", Contents(kind, nativeHeader, baseId, nativeSchema, pageRoles)),
            new OfficeCompoundStream("Quill/QuillSub/CONTENTS", QuillCandidate.Create(text)),
            new OfficeCompoundStream("Escher/EscherStm", OfficeArt(kind, image, baseId + 2)),
            new OfficeCompoundStream("Escher/EscherDelayStm", Array.Empty<byte>())
        };
        if (bootstrapMetadata) streams.Add(new OfficeCompoundStream("\u0001CompObj",
            OfficeOleCompoundObjectWriter.Write(rootClassId ?? throw new ArgumentException("Bootstrap metadata requires its native class ID."),
                "Microsoft Publisher 3.0", "MSPublisher.3", string.Empty)));
        return OfficeCompoundFileWriter.Write(streams, rootClassId);
    }

    private static byte[] Contents(string kind, bool nativeHeader, uint baseId, bool nativeSchema, bool pageRoles) {
        bool hasShape = kind != "blank";
        byte[] dimensions = Variable(0x12, 0x88, Scalar(1, 0x20, WidthEmu), Scalar(2, 0x20, HeightEmu));
        byte[] pages = Variable(2, nativeHeader ? (byte)0xA0 : (byte)0x90, Scalar(0, 0x70, baseId + 1));
        byte[] document = pageRoles
            ? Chunk(Scalar(1, 0x20, 1), pages, dimensions,
                ShortScalar(0x2C, 0x18, 1), ShortScalar(0x2D, 0x18, 0))
            : Chunk(dimensions, pages);
        var chunks = new List<(uint Kind, uint? Parent, byte[] Bytes)> {
            (0x44, null, document),
            (0x43, baseId, Chunk(Scalar(1, 0x20, hasShape ? 1U : 0U),
                Variable(2, nativeHeader ? (byte)0xA0 : (byte)0x90, hasShape ? Scalar(0, 0x70, baseId + 2) : Array.Empty<byte>())))
        };
        if (hasShape) chunks.Add((kind == "image" ? 0x20U : 1U, baseId + 1,
            kind == "text" ? Chunk(Scalar(0x27, 0x20, 1), Scalar(0x28, 0x20, 0)) : Chunk()));
        byte[] header = nativeSchema ? NativeSchemaPrefix.Create(kind, baseId, 0, 0) : new byte[32];
        var offsets = new List<int>();
        int offset = header.Length;
        foreach (var chunk in chunks) { offsets.Add(offset); offset += chunk.Bytes.Length; }
        int trailer = offset;
        var slots = Enumerable.Range(0, checked((int)baseId)).Select(_ => Marker(0, 0x78)).ToList();
        slots.AddRange(chunks.Select((chunk, index) =>
            Variable(0, 0x88,
                nativeHeader ? ShortScalar(2, 0x18, checked((ushort)chunk.Kind)) : Scalar(2, 0x20, chunk.Kind),
                Scalar(4, nativeHeader ? (byte)0xB8 : (byte)0x20, (uint)offsets[index]),
                chunk.Parent.HasValue ? Scalar(5, nativeHeader ? (byte)0x68 : (byte)0x70, chunk.Parent.Value) : Array.Empty<byte>())));
        byte[] directory = Variable(nativeHeader ? (byte)3 : (byte)1, 0x90, slots.ToArray());
        header[0] = 0xE8; header[1] = 0xAC; header[2] = 0x2C;
        BitConverter.GetBytes(trailer).CopyTo(header, 0x1A);
        byte[] bytes = Join(header, Join(chunks.Select(chunk => chunk.Bytes).ToArray()), nativeHeader
            ? Chunk(Scalar(1, 0x20, baseId + (uint)chunks.Count), Scalar(2, 0x20, baseId + (uint)chunks.Count - 1), directory)
            : Chunk(Scalar(0, 0x20, 0), directory, Scalar(2, 0x20, 0)));
        if (nativeSchema) {
            byte[] prefix = NativeSchemaPrefix.Create(kind, baseId, bytes.Length, trailer);
            if (prefix.Length != header.Length) throw new InvalidDataException("Native schema fixup changed its declared prefix size.");
            prefix.CopyTo(bytes, 0);
        } else if (nativeHeader) {
            // These fields are independently observed in the controlled native
            // publications. Auxiliary bootstrap records remain unresolved.
            BitConverter.GetBytes(0x000A03E8U).CopyTo(bytes, 4);
            BitConverter.GetBytes(bytes.Length).CopyTo(bytes, 8);
            BitConverter.GetBytes(0x0100001AU).CopyTo(bytes, 12);
        }
        return bytes;
    }

    private static byte[] OfficeArt(string kind, byte[]? image, uint nativeShapeId) {
        if (kind == "blank") return Art(15, 0xF000);
        var properties = new List<(ushort Id, uint Value)> {
            (0x181, 0x0000FF), // Solid red in the common OfficeArt RGB encoding.
            (0x1BF, 0x00100010),
            (0x1FF, 0x00080000) // No line.
        };
        if (kind == "text") {
            properties[1] = (0x1BF, 0x00100000); // Text frame has no fill.
            properties.Add((0x81, 0)); properties.Add((0x82, 0));
            properties.Add((0x83, 0)); properties.Add((0x84, 0));
        }
        if (image != null) properties.Add((0x4104, 1)); // Picture index carries the native BLIP-ID flag.
        byte[] fopt = Bytes(writer => {
            foreach (var property in properties.OrderBy(property => property.Id & 0x3FFF)) {
                writer.Write(property.Id); writer.Write(property.Value);
            }
        });
        uint Coord(double points, uint pageSize) => unchecked((uint)(int)(points * 12700 - pageSize / 2D));
        byte[] shape = Art(15, 0xF004,
            Art((ushort)(((kind == "image" ? 75 : 1) << 4) | 2), 0xF00A,
                Bytes(writer => { writer.Write(1025U); writer.Write(0xA00U); })),
            Art((ushort)((properties.Count << 4) | 3), 0xF00B, fopt),
            Art(0, 0xF010, Client((0x2001, Coord(72, WidthEmu)), (0x2002, Coord(72, HeightEmu)),
                (0x2003, Coord(360, WidthEmu)), (0x2004, Coord(180, HeightEmu)))),
            Art(0, 0xF011, Client((0x6801, nativeShapeId))));
        byte[] root = Art(15, 0xF004,
            Art(1, 0xF009, new byte[16]),
            Art(2, 0xF00A, Bytes(writer => { writer.Write(1024U); writer.Write(5U); })));
        byte[] drawing = Art(15, 0xF002,
            Art(16, 0xF008, Bytes(writer => { writer.Write(2U); writer.Write(1025U); })),
            Art(15, 0xF003, root, shape));
        byte[] dgg = image == null ? Art(15, 0xF000) : Art(15, 0xF000,
            Art(31, 0xF001, OfficeArtBlipStoreEntryWriter.CreateEmbedded(image,
                image.Length >= 8 && image[0] == 0x89 && image[1] == 0x50 ? "image/png" : "image/jpeg")));
        return Join(dgg, new byte[4], drawing);
    }
}
