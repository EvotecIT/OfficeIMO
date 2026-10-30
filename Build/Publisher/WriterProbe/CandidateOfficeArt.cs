using OfficeIMO.Drawing.Binary;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;
using static OfficeIMO.Publisher.WriterProbe.NativePublicationGraph;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Constructs the probe's calculated, format-owned OfficeArt drawing topology.</summary>
internal static class CandidateOfficeArt {
    internal static byte[] Create(Graph graph) {
        byte[]? image = Image(graph);
        uint shapeCount = (uint)graph.Objects.Count + 1;
        uint lastShapeId = RootShapeId + shapeCount - 1;
        byte[] fdgg = Art(0, 0xF006, Bytes(writer => {
            // Follow the native/PPT allocation convention: one beyond the
            // greatest allocated SPID. No hidden background shape is emitted.
            writer.Write(lastShapeId + 1); // spidMax
            writer.Write(2U); // cidcl: one IDCL entry plus one.
            writer.Write(shapeCount); // cspSaved
            writer.Write(1U); // cdgSaved
            writer.Write(DrawingId); writer.Write(shapeCount); // IDCL
        }));
        var groupRecords = new List<byte[]> { fdgg };
        if (image != null) groupRecords.Add(Art(31, 0xF001,
            OfficeArtBlipStoreEntryWriter.CreateDelayed(image, ContentType(image), delayedStreamOffset: 0)));
        byte[] dgg = Art(15, 0xF000, groupRecords.ToArray());
        byte[] root = Art(15, 0xF004,
            Art(1, 0xF009, new byte[16]),
            Art(2, 0xF00A, Bytes(writer => { writer.Write(RootShapeId); writer.Write(5U); })));
        byte[][] shapes = new[] { root }.Concat(graph.Objects.Select(item => Leaf(item, graph.Source.WidthEmu, graph.Source.HeightEmu))).ToArray();
        byte[] drawing = Art(15, 0xF002,
            Art((ushort)(DrawingId << 4), 0xF008, Bytes(writer => { writer.Write(shapeCount); writer.Write(lastShapeId); })),
            Art(15, 0xF003, shapes));
        return Join(dgg, new byte[4], drawing);
    }

    internal static byte[] CreateImageDelayStream(Graph graph) {
        byte[]? image = Image(graph);
        return image == null ? Array.Empty<byte>() : OfficeArtBlipStoreEntryWriter.CreateBlipRecord(image, ContentType(image));
    }

    private static byte[]? Image(Graph graph) => graph.Objects.SingleOrDefault(item => item.Source.Image != null)?.Source.Image;

    private static string ContentType(byte[] image) =>
        image.Length >= 8 && image[0] == 0x89 && image[1] == 0x50 ? "image/png" : "image/jpeg";

    private static byte[] Leaf(ObjectBinding item, uint width, uint height) {
        bool isTextBox = item.Source.Kind == CandidateObjectKind.TextBox;
        bool hasImage = item.Source.Image != null;
        var properties = new List<(ushort Id, uint Value)> {
            (0x181, 0x0000FF), // Solid red in the common OfficeArt RGB encoding.
            (0x1BF, 0x00100010),
            (0x1FF, 0x00080000) // No line.
        };
        if (isTextBox) {
            properties[1] = (0x1BF, 0x00100000); // Text frame has no fill.
            properties.Add((0x81, 0)); properties.Add((0x82, 0));
            properties.Add((0x83, 0)); properties.Add((0x84, 0));
        }
        if (hasImage) properties.Add((0x4104, 1)); // Picture index carries the native BLIP-ID flag.
        byte[] fopt = Bytes(writer => {
            foreach (var property in properties.OrderBy(property => property.Id & 0x3FFF)) {
                writer.Write(property.Id); writer.Write(property.Value);
            }
        });
        uint Coord(double points, uint pageSize) => unchecked((uint)(int)(points * 12700 - pageSize / 2D));
        return Art(15, 0xF004,
            Art((ushort)(((item.Source.Kind == CandidateObjectKind.Picture ? 75 : isTextBox ? 202 : 1) << 4) | 2), 0xF00A,
                Bytes(writer => { writer.Write(item.Spid); writer.Write(0xA00U); })),
            Art((ushort)((properties.Count << 4) | 3), 0xF00B, fopt),
            Art(0, 0xF010, Client((0x2001, Coord(item.Source.Placement.Left, width)), (0x2002, Coord(item.Source.Placement.Top, height)),
                (0x2003, Coord(item.Source.Placement.Left + item.Source.Placement.Width, width)),
                (0x2004, Coord(item.Source.Placement.Top + item.Source.Placement.Height, height)))),
            Art(0, 0xF011, Client((0x6801, item.Id))),
            // Publisher's native text-bearing shapes carry this ClientTextbox
            // OPL with their story key and the observed client record header.
            item.Story != null ? Art(0x01AA, 0xF00D, Client((0x2001, item.Story.Key))) : Array.Empty<byte>());
    }
}
