using OfficeIMO.Drawing.Binary;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;
using static OfficeIMO.Publisher.WriterProbe.NativePublicationGraph;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Constructs the probe's calculated, format-owned OfficeArt drawing topology.</summary>
internal static class CandidateOfficeArt {
    internal static IReadOnlyList<ImageBinding> BindImages(Graph graph) {
        var images = new List<ImageBinding>();
        var assets = new Dictionary<int, ImageBinding>();
        uint offset = 0;
        foreach (ObjectBinding item in graph.Objects.Where(item => item.Source.Image != null)) {
            byte[] image = item.Source.Image!;
            if (item.Source.ImageId is { } imageId && assets.TryGetValue(imageId, out ImageBinding? existing)) {
                if (!existing.Image.SequenceEqual(image)) throw new InvalidDataException("One model image identifier has conflicting payloads.");
                existing.ObjectIds.Add(item.Id);
                continue;
            }
            byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(image, ContentType(image));
            var binding = new ImageBinding(new List<uint> { item.Id }, checked((uint)images.Count + 1), offset, image, blip);
            images.Add(binding);
            if (item.Source.ImageId is { } id) assets.Add(id, binding);
            offset = checked(offset + (uint)blip.Length);
        }
        if (images.Count > 4095) throw new NotSupportedException("The OfficeArt store supports at most 4,095 image entries.");
        return images;
    }

    internal static byte[] Create(Graph graph, IReadOnlyList<ImageBinding> images) {
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
        // Native references follow model asset identity, never byte equality.
        if (images.Count != 0) groupRecords.Add(Art(checked((ushort)((images.Count << 4) | 15)), 0xF001,
            images.Select(image => OfficeArtBlipStoreEntryWriter.CreateDelayed(image.Image,
                ContentType(image.Image), image.Offset, referenceCount: checked((uint)image.ObjectIds.Count))).ToArray()));
        byte[] dgg = Art(15, 0xF000, groupRecords.ToArray());
        byte[] root = Art(15, 0xF004,
            Art(1, 0xF009, new byte[16]),
            Art(2, 0xF00A, Bytes(writer => { writer.Write(RootShapeId); writer.Write(5U); })));
        var imageIndices = images.SelectMany(image => image.ObjectIds.Select(id => (Id: id, image.Index)))
            .ToDictionary(item => item.Id, item => item.Index);
        byte[][] shapes = new[] { root }.Concat(graph.Objects.Select(item => Leaf(item, graph.Source.WidthEmu,
            graph.Source.HeightEmu, imageIndices.TryGetValue(item.Id, out uint index) ? index : null))).ToArray();
        byte[] drawing = Art(15, 0xF002,
            Art((ushort)(DrawingId << 4), 0xF008, Bytes(writer => { writer.Write(shapeCount); writer.Write(lastShapeId); })),
            Art(15, 0xF003, shapes));
        return Join(dgg, new byte[4], drawing);
    }

    internal static byte[] CreateImageDelayStream(IReadOnlyList<ImageBinding> images) => Join(images.Select(image => image.Blip).ToArray());

    internal sealed record ImageBinding(List<uint> ObjectIds, uint Index, uint Offset, byte[] Image, byte[] Blip);

    private static string ContentType(byte[] image) =>
        image.Length >= 8 && image[0] == 0x89 && image[1] == 0x50 ? "image/png" : "image/jpeg";

    private static byte[] Leaf(ObjectBinding item, uint width, uint height, uint? imageIndex) {
        bool isTextBox = item.Source.Kind == CandidateObjectKind.TextBox;
        bool isTable = item.Source.Kind == CandidateObjectKind.Table;
        var properties = new List<(ushort Id, uint Value)> {
            (0x181, 0x0000FF), // Solid red in the common OfficeArt RGB encoding.
            (0x1BF, 0x00100010),
            (0x1FF, 0x00080000) // No line.
        };
        if (item.Source.Paint is { } paint) {
            if (paint.StrokeColor.HasValue || (paint.FillOpacity ?? 1) != 1 || (paint.StrokeOpacity ?? 1) != 1
                || paint.FillColor is { A: not 255 })
                throw new NotSupportedException("The model specimen supports opaque solid fills or no fill, without an outline.");
            OfficeIMO.Drawing.OfficeColor? fill = paint.FillColor;
            properties[0] = (0x181, fill.HasValue ? (uint)(fill.Value.R | fill.Value.G << 8 | fill.Value.B << 16) : 0);
            properties[1] = (0x1BF, fill.HasValue ? 0x00100010U : 0x00100000U);
        }
        if (isTextBox || isTable) {
            properties[1] = (0x1BF, 0x00100000); // Text frame has no fill.
            properties.Add((0x81, 0)); properties.Add((0x82, 0));
            properties.Add((0x83, 0)); properties.Add((0x84, 0));
            if (item.Source.ColumnSpacing.HasValue)
                properties.Add((0x8D, checked((uint)(item.Source.ColumnSpacing.Value * 12700))));
        }
        if (imageIndex.HasValue) properties.Add((0x4104, imageIndex.Value)); // Native BLIP-ID flag.
        if (item.Source.Crop.HasCrop)
            properties.AddRange(OfficeArtPictureProperties.EncodeCrop(item.Source.Crop, out _)
                .Select(property => (property.RawOperationId, property.Value)));
        byte[] fopt = Bytes(writer => {
            foreach (var property in properties.OrderBy(property => property.Id & 0x3FFF)) {
                writer.Write(property.Id); writer.Write(property.Value);
            }
        });
        ushort propertyCount = checked((ushort)properties.Count);
        if (item.Source.ImageName is { } name) {
            byte[] text = System.Text.Encoding.Unicode.GetBytes(Path.GetFileNameWithoutExtension(name) + "\0");
            fopt = OfficeArtPropertyTableEditor.Rewrite(fopt, 0, fopt.Length, propertyCount,
                new Dictionary<ushort, OfficeArtProperty> {
                    [0x105] = new(0, 0xC105, checked((uint)text.Length), text.Length, complexData: text)
                }, out propertyCount) ?? throw new InvalidDataException("The picture filename could not be encoded.");
        }
        uint Coord(double points, uint pageSize) => unchecked((uint)(int)(points * 12700 - pageSize / 2D));
        return Art(15, 0xF004,
            Art((ushort)(((item.Source.Kind == CandidateObjectKind.Picture ? 75 : isTable ? 201 : isTextBox ? 202 : 1) << 4) | 2), 0xF00A,
                Bytes(writer => { writer.Write(item.Spid); writer.Write(0xA00U); })),
            Art((ushort)((propertyCount << 4) | 3), 0xF00B, fopt),
            Art(0, 0xF010, Client((0x2001, Coord(item.Source.Placement.Left, width)), (0x2002, Coord(item.Source.Placement.Top, height)),
                (0x2003, Coord(item.Source.Placement.Left + item.Source.Placement.Width, width)),
                (0x2004, Coord(item.Source.Placement.Top + item.Source.Placement.Height, height)))),
            Art(0, 0xF011, Client((0x6801, item.Id))),
            // Publisher's native text-bearing shapes carry this ClientTextbox
            // OPL with their story key and the observed client record header.
            item.Story != null ? Art(0x01AA, 0xF00D, Client((0x2001, item.Story.Key))) : Array.Empty<byte>());
    }
}
