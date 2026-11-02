using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherEscherTableUtilityPlan {
    private static SourceProfile? ReadProfile(PublisherBinaryData data, PublisherEscherReader reader, PublisherParseContext context) {
        IReadOnlyList<PublisherEscherRecord> top = reader.Records(0, data.Length);
        if (top.Count is not (2 or 4) || top[0].Kind != 0xF000 || (top[0].Initial & 15) != 15) return null;
        bool emptyScopes = top.Count == 4;
        var globals = reader.Records(top[0].Offset, top[0].End);
        if (globals.Count(item => item.Kind == 0xF006) != 1
            || globals.Any(item => item.Kind is not (0xF006 or 0xF001 or 0xF00B or 0xF122 or 0xF11E))
            || !QualifiedDefaults(data, globals, context)) return null;
        PublisherEscherRecord allocation = globals.Single(item => item.Kind == 0xF006);
        if (allocation.Initial != 0 || allocation.Length != 24 || data.U32(allocation.Offset + 4) != 2
            || data.U32(allocation.Offset + 12) != (emptyScopes ? 3U : 1U)
            || data.U32(allocation.Offset + 16) != 1) return null;
        var ids = new HashSet<uint>();
        PublisherEscherRecord ordinary = top[top.Count - 1];
        DrawingProfile? normal = ReadDrawing(ordinary, 1, 0, background: false);
        if (normal == null || normal.Shapes.Count < 2) return null;
        uint maximum = data.U32(allocation.Offset);
        if (data.U32(allocation.Offset + 8) != normal.Shapes.Count + (emptyScopes ? 2 : 0)
            || data.U32(allocation.Offset + 20) != (normal.Cursor & 1023) + 1
            || maximum <= normal.Cursor || maximum > 2048) return null;
        DrawingProfile? utility = null; uint auxiliaryUsed = 1;
        if (emptyScopes) {
            utility = ReadDrawing(top[1], 3, 2, background: false);
            DrawingProfile? auxiliary = ReadDrawing(top[2], 2, 1, background: true);
            if (utility == null || auxiliary == null || utility.Shapes.Count != 1 || auxiliary.Shapes.Count != 1
                || utility.Cursor != 3072 || auxiliary.Cursor != 2049) return null;
            auxiliaryUsed = 2;
        }
        return new(allocation, ordinary, new HashSet<int>(normal.Shapes.Select(item => item.Offset)), normal.Shapes.Count, auxiliaryUsed, utility);

        DrawingProfile? ReadDrawing(PublisherEscherRecord record, uint id, uint instance, bool background) {
            if (record.Kind != 0xF002 || (record.Initial & 15) != 15 || record.Offset < 12 || data.U32(record.Offset - 12) != instance) return null;
            var children = reader.Records(record.Offset, record.End);
            if (children.Count != (background ? 3 : 2) || children.Count(item => item.Kind == 0xF008) != 1
                || children.Count(item => item.Kind == 0xF003) != 1 || children.Count(item => item.Kind == 0xF004) != (background ? 1 : 0)) return null;
            var fdg = children.Single(item => item.Kind == 0xF008); var group = children.Single(item => item.Kind == 0xF003);
            if (fdg.Initial != id << 4 || fdg.Length != 8 || (group.Initial & 15) != 15) return null;
            var shapes = reader.Records(group.Offset, group.End); uint cursor = data.U32(fdg.Offset + 4);
            if (shapes.Count < 1 || shapes.Count > 1024 || shapes.Any(item => item.Kind != 0xF004)
                || data.U32(fdg.Offset) != shapes.Count || cursor < id << 10 || cursor >= (id + 1) << 10) return null;
            for (int index = 0; index < shapes.Count; index++) {
                context.Token.ThrowIfCancellationRequested();
                if (!Shape(shapes[index], index == 0 ? 5U : 0xA00U, index == 0)) return null;
            }
            if (background && !Shape(children.Single(item => item.Kind == 0xF004), 0xC00, root: false)) return null;
            return new(fdg, group, shapes, cursor);

            bool Shape(PublisherEscherRecord shape, uint flags, bool root) {
                var values = reader.Records(shape.Offset, shape.End);
                var fsps = values.Where(item => item.Kind == 0xF00A).ToArray();
                if (fsps.Length != 1 || fsps[0].Length != 8) return false;
                var fsp = fsps[0]; uint spid = data.U32(fsp.Offset);
                if (spid < id << 10 || spid > cursor || !ids.Add(spid) || data.U32(fsp.Offset + 4) != flags) return false;
                if (!root) return true;
                var spaces = values.Where(item => item.Kind == 0xF009).ToArray();
                return spid == id << 10 && fsp.Initial == 2 && values.Count == 2 && spaces.Length == 1 && spaces[0].Length == 16
                    && Enumerable.Range(0, 4).All(axis => data.U32(spaces[0].Offset + axis * 4) == 0);
            }
        }
    }

    private static bool QualifiedDefaults(PublisherBinaryData data, IReadOnlyList<PublisherEscherRecord> records, PublisherParseContext context) {
        var properties = new List<OfficeArtProperty>();
        foreach (var record in records.Where(item => item.Kind is 0xF00B or 0xF122)) {
            if ((record.Initial & 15) != 3 || (record.Initial >> 4) * 6 > record.Length) return false;
            foreach (var property in OfficeArtPropertyTableReader.Read(data.Bytes, record.Offset, record.Length, (ushort)(record.Initial >> 4))) {
                context.Record();
                if (property.IsComplex || property.IsBlipId) return false;
                properties.Add(property);
            }
        }
        // Fresh utilities own solid color/fill/line bits, while inherited
        // opacity, type, hidden state, shadow and transform must stay simple.
        properties.Add(new(0, 0x0181, 0x00FFFFFF));
        properties.Add(new(0, 0x01BF, 0x001F001C));
        properties.Add(new(0, 0x01FF, 0x00080000));
        var style = OfficeArtShapeStyle.Decode(properties);
        var transform = OfficeArtShapeTransform.Decode(0xA00, properties, relativeToParent: false);
        return style.FillEnabled != false && style.FillType.GetValueOrDefault() == 0 && style.FillOpacity.GetValueOrDefault(1) == 1
            && style.Hidden != true && style.ShadowEnabled != true && transform.RotationDegrees.GetValueOrDefault() == 0
            && !transform.FlipHorizontal && !transform.FlipVertical;
    }
}
