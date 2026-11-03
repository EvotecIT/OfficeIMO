using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.Internal.PublisherNativeCreationEncoding;
using static OfficeIMO.Publisher.Internal.PublisherNativeCreationGraph;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Constructs the probe's calculated, format-owned OfficeArt drawing topology.</summary>
internal static partial class PublisherNativeCreationArtwork {
    internal const int MaximumOrdinaryShapes = 1024;
    internal const int MaximumImages = 4095;

    internal static IReadOnlyList<ImageBinding> BindImages(Graph graph, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        var images = new List<ImageBinding>();
        var assets = new Dictionary<int, ImageBinding>();
        uint offset = 0;
        foreach (ObjectBinding item in graph.Objects) {
            context.Token.ThrowIfCancellationRequested();
            if (item.Source.Image is not { } image) continue;
            if (item.Source.ImageId is { } imageId && assets.TryGetValue(imageId, out ImageBinding? existing)) {
                if (!ImagePayloadsEqual(existing.Image, image, context.Token)) throw new InvalidDataException("One model image identifier has conflicting payloads.");
                existing.ObjectIds.Add(item.Id);
                continue;
            }
            context.Token.ThrowIfCancellationRequested();
            byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(image, ContentType(image));
            context.Token.ThrowIfCancellationRequested();
            var binding = new ImageBinding(new List<uint> { item.Id }, checked((uint)images.Count + 1), offset, image, blip);
            images.Add(binding);
            if (item.Source.ImageId is { } id) assets.Add(id, binding);
            offset = checked(offset + (uint)blip.Length);
        }
        if (images.Count > MaximumImages) throw new NotSupportedException("The OfficeArt store supports at most 4,095 image entries.");
        return images;
    }

    internal static byte[] Create(Graph graph, IReadOnlyList<ImageBinding> images, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        if (graph.Source.WidthEmu is 0 or > int.MaxValue || graph.Source.HeightEmu is 0 or > int.MaxValue)
            throw new NotSupportedException("Native page dimensions require positive signed-EMU values.");
        uint shapeCount = (uint)graph.Objects.Count + 1;
        uint lastShapeId = RootShapeId + shapeCount - 1;
        var allocation = PublisherEscherDrawingAllocation.Fresh(1, context.Token);
        if (shapeCount > MaximumOrdinaryShapes || !allocation.StartDrawing(DrawingId, shapeCount, context.Token))
            throw new NotSupportedException("The ordinary object drawing specimen is limited to one SPID cluster; table utility drawings can span clusters.");
        var tableDrawing = TableDrawing(graph, allocation, context);
        // The shared allocator owns every IDCL entry and the exclusive maximum;
        // table utility rollover does not add another saved drawing.
        byte[] fdgg = allocation.Record(shapeCount + (tableDrawing != null ? tableDrawing.Count + 1 : 0),
            tableDrawing != null ? 3U : 1U, context.Token);
        var groupRecords = new List<byte[]> { fdgg };
        // Native references follow model asset identity, never byte equality.
        var imageRecords = new byte[images.Count][];
        var imageIndices = new Dictionary<uint, uint>();
        for (int index = 0; index < images.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            ImageBinding image = images[index];
            imageRecords[index] = OfficeArtBlipStoreEntryWriter.CreateDelayed(image.Image,
                ContentType(image.Image), image.Offset, referenceCount: checked((uint)image.ObjectIds.Count));
            context.Token.ThrowIfCancellationRequested();
            foreach (uint id in image.ObjectIds) {
                context.Token.ThrowIfCancellationRequested();
                imageIndices.Add(id, image.Index);
            }
        }
        if (images.Count != 0) groupRecords.Add(Art(checked((ushort)((images.Count << 4) | 15)), 0xF001, imageRecords));
        byte[] dgg = Art(15, 0xF000, groupRecords.ToArray());
        var shapes = new byte[graph.Objects.Count + 1][];
        shapes[0] = DrawingRoot(DrawingId);
        for (int index = 0; index < graph.Objects.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            ObjectBinding item = graph.Objects[index];
            shapes[index + 1] = Leaf(item, graph.Source.WidthEmu, graph.Source.HeightEmu,
                imageIndices.TryGetValue(item.Id, out uint imageIndex) ? imageIndex : null, context);
        }
        byte[] drawing = Drawing(DrawingId, shapes, lastShapeId);
        // The four-byte prefixes select drawing instances, rather than padding.
        // Native table controls serialize utility instance 2, empty instance 1,
        // then ordinary instance 0; SPID drawing identifiers remain 3, 2 and 1.
        context.Token.ThrowIfCancellationRequested();
        byte[] result = tableDrawing != null ? Join(dgg, Bytes(writer => writer.Write(2U)), tableDrawing.Bytes,
            Bytes(writer => writer.Write(1U)), PublisherTableUtilityDrawingEncoding.EmptyDrawing(2), new byte[4], drawing)
            : Join(dgg, new byte[4], drawing);
        context.Token.ThrowIfCancellationRequested();
        return result;
    }

    private static byte[] DrawingRoot(uint id) => Art(15, 0xF004, Art(1, 0xF009, new byte[16]),
        Art(2, 0xF00A, Bytes(writer => { writer.Write(id << 10); writer.Write(5U); })));

    private static byte[] Drawing(uint id, byte[][] shapes, uint lastSpid) => Art(15, 0xF002,
        Art(checked((ushort)(id << 4)), 0xF008, Bytes(writer => { writer.Write((uint)shapes.Length); writer.Write(lastSpid); })),
        Art(15, 0xF003, shapes));

    internal static byte[] CreateImageDelayStream(IReadOnlyList<ImageBinding> images, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        byte[] result = Bytes(writer => {
            foreach (ImageBinding image in images) {
                context.Token.ThrowIfCancellationRequested();
                writer.Write(image.Blip);
            }
        });
        context.Token.ThrowIfCancellationRequested();
        return result;
    }

    internal sealed record ImageBinding(List<uint> ObjectIds, uint Index, uint Offset, byte[] Image, byte[] Blip);

    private static string ContentType(byte[] image) =>
        image.Length >= 8 && image[0] == 0x89 && image[1] == 0x50 ? "image/png" : "image/jpeg";

    /// <summary>Compares shared asset identity without leaving large payload comparisons uncancellable.</summary>
    internal static bool ImagePayloadsEqual(byte[] first, byte[] second, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (first.Length != second.Length) return false;
        for (int index = 0; index < first.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            if (first[index] != second[index]) return false;
        }
        token.ThrowIfCancellationRequested();
        return true;
    }

    private static byte[] Leaf(ObjectBinding item, uint width, uint height, uint? imageIndex, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        bool isTextBox = item.Source.Kind == PublisherNativeCreationObjectKind.TextBox;
        bool isTable = item.Source.Kind == PublisherNativeCreationObjectKind.Table;
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
            uint Distance(double points) {
                if (points < 0 || !PublisherNativeGeometry.Emu(points, 0, out int encoded) || encoded < 0)
                    throw new NotSupportedException("Native frame spacing and insets require exact nonnegative signed-EMU values.");
                return (uint)encoded;
            }
            if (item.Source.FrameLayout is { } layout) {
                properties.Add((0x81, Distance(layout.Padding.Left))); properties.Add((0x82, Distance(layout.Padding.Top)));
                properties.Add((0x83, Distance(layout.Padding.Right))); properties.Add((0x84, Distance(layout.Padding.Bottom)));
                properties.Add((0x8C, checked((uint)layout.ColumnCount)));
                properties.Add((0x8D, Distance(layout.ColumnSpacing)));
            } else {
                properties.Add((0x81, 0)); properties.Add((0x82, 0));
                properties.Add((0x83, 0)); properties.Add((0x84, 0));
            }
            if (item.Source.FrameLayout == null && item.Source.ColumnSpacing.HasValue)
                properties.Add((0x8D, Distance(item.Source.ColumnSpacing.Value)));
        }
        if (imageIndex.HasValue) properties.Add((0x4104, imageIndex.Value)); // Native BLIP-ID flag.
        if (item.Source.Crop.HasCrop)
            properties.AddRange(OfficeArtPictureProperties.EncodeCrop(item.Source.Crop, out _)
                .Select(property => (property.RawOperationId, property.Value)));
        byte[] fopt = Bytes(writer => {
            foreach (var property in properties.OrderBy(property => property.Id & 0x3FFF)) {
                context.Token.ThrowIfCancellationRequested();
                writer.Write(property.Id); writer.Write(property.Value);
            }
        });
        ushort propertyCount = checked((ushort)properties.Count);
        if (item.Source.ImageName is { } name) {
            context.Token.ThrowIfCancellationRequested();
            byte[] text = System.Text.Encoding.Unicode.GetBytes(Path.GetFileNameWithoutExtension(name) + "\0");
            fopt = OfficeArtPropertyTableEditor.Rewrite(fopt, 0, fopt.Length, propertyCount,
                new Dictionary<ushort, OfficeArtProperty> {
                    [0x105] = new(0, 0xC105, checked((uint)text.Length), text.Length, complexData: text)
                }, out propertyCount) ?? throw new InvalidDataException("The picture filename could not be encoded.");
            context.Token.ThrowIfCancellationRequested();
        }
        uint Coord(double points, double origin) {
            if (!PublisherNativeGeometry.Emu(points, origin / 12700D, out int encoded))
                throw new NotSupportedException("Native object bounds require complete exact signed-EMU coordinates.");
            return unchecked((uint)encoded);
        }
        return Art(15, 0xF004,
            Art((ushort)(((item.Source.Kind == PublisherNativeCreationObjectKind.Picture ? 75 : isTable ? 201 : isTextBox ? 202 : 1) << 4) | 2), 0xF00A,
                Bytes(writer => { writer.Write(item.Spid); writer.Write(0xA00U); })),
            Art((ushort)((propertyCount << 4) | 3), 0xF00B, fopt),
            Art(0, 0xF010, Client((0x2001, Coord(item.Source.Placement.Left, item.HorizontalOriginEmu)), (0x2002, Coord(item.Source.Placement.Top, height / 2D)),
                (0x2003, Coord(item.Source.Placement.Left + item.Source.Placement.Width, item.HorizontalOriginEmu)),
                (0x2004, Coord(item.Source.Placement.Top + item.Source.Placement.Height, height / 2D)))),
            Art(0, 0xF011, Client((0x6801, item.Id))),
            // Publisher's native text-bearing shapes carry this ClientTextbox
            // OPL with their story key and the observed client record header.
            item.Story != null ? Art(0x01AA, 0xF00D, Client((0x2001, item.Story.Key))) : Array.Empty<byte>());
    }
}
