using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly List<PublisherNativeStreamEdit> _tableTrackEdits = new();
    private bool _hasTableBounds;

    /// <summary>Moves direct table frames, scales a fixed grid or updates existing utility paint without changing cell ownership.</summary>
    private void CompareTable(PublisherTableState before, PublisherTableState after) {
        _token.ThrowIfCancellationRequested();
        if (TableEqual(before, after)) return;
        OfficeShape? paint = before.CreatePaint();
        if (before.Id != after.Id || before.StoryId != after.StoryId || before.PageId != after.PageId
            || !before.HasPlacement || !after.HasPlacement || before.Hidden != after.Hidden
            || !before.Padding.Equals(after.Padding) || before.IsAuthored != after.IsAuthored
            || before.BoundsEditLimitation != null || after.BoundsEditLimitation != null
            || before.PageTransform != OfficeTransform.Translate(before.X, before.Y)
            || after.PageTransform != OfficeTransform.Translate(after.X, after.Y)
            || !before.Cells.SequenceEqual(after.Cells)
            || !ShapeEqual(paint == null ? null : PublisherShapeGeometry.Resize(paint, after.Width, after.Height), after.CreatePaint())
            || !ScaledTracks(before.Columns, after.Columns, before.Width, after.Width)
            || !ScaledTracks(before.Rows, after.Rows, before.Height, after.Height)) {
            Reject("Contents/object/" + before.Id, "Table changes extend beyond untransformed placement, proportional fixed-grid resizing and existing utility paint."); return;
        }
        CompareTableDecoration(before, after);
        if ((before.Width != after.Width || before.Height != after.Height) && !PlanTableTracks(before, after)) {
            Reject("Contents/object/" + before.Id, "Table resizing requires complete direct fixed tracks, matching nominal and laid-out sizes, and exact positive signed-EMU dimensions."); return;
        }
        if (before.X != after.X || before.Y != after.Y || before.Width != after.Width || before.Height != after.Height) {
            Bounds(before.Id, before.PageId, after.X, after.Y, after.Width, after.Height);
            _hasTableBounds = true;
        }
    }

    private static bool ScaledTracks(IReadOnlyList<double> before, IReadOnlyList<double> after, double oldExtent, double newExtent) {
        if (before.Count != after.Count) return false;
        for (int i = 0; i < before.Count; i++) {
            // Repeated public resizes can reach the same native track through
            // slightly different floating-point arithmetic paths.
            if (!Emu(before[i] / oldExtent * newExtent, 0, out int expected)
                || !Emu(after[i], 0, out int actual) || expected != actual) return false;
        }
        return true;
    }

    /// <summary>Plans fixed-size Contents changes. Expanded native rows retain their distinct nominal sizes and are refused.</summary>
    private bool PlanTableTracks(PublisherTableState before, PublisherTableState after) {
        if (!_imported.Contents.Shapes.TryGetValue(before.Id, out PublisherSourceShape? shape) || shape.Chunk.Kind != 0x10
            || shape.Table == null || !Emu(before.Width, 0, out int oldWidth) || !Emu(before.Height, 0, out int oldHeight)
            || !Emu(after.Width, 0, out int width) || width <= 0 || !Emu(after.Height, 0, out int height) || height <= 0) return false;
        var data = new PublisherBinaryData(_imported.NativeSource!.ContentsBytes, "Contents");
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = _sourceLimits }, _token);
        var reader = new PublisherBlockReader(data, context);
        var edits = new List<PublisherNativeStreamEdit>();
        if (!Extent(0x68, oldWidth, width) || !Extent(0x69, oldHeight, height)) return false;
        PublisherBlock[] dimensions = shape.Fields.Where(field => field.Id == 0x6D).ToArray();
        if (dimensions.Length != 1 || dimensions[0].Type != 0x90) return false;
        IReadOnlyList<PublisherBlock> tracks = reader.Children(dimensions[0]);
        if (tracks.Count != before.Columns.Count + before.Rows.Count) return false;
        if (!Axis(before.Columns, after.Columns, 0, oldWidth, width)
            || !Axis(before.Rows, after.Rows, before.Columns.Count, oldHeight, height)) return false;
        _tableTrackEdits.AddRange(edits);
        return true;

        bool Extent(byte id, int previous, int value) {
            PublisherBlock[] fields = shape.Fields.Where(field => field.Id == id).ToArray();
            return fields.Length == 1 && Scalar(fields[0], previous, value);
        }
        bool Scalar(PublisherBlock field, int previous, int value) {
            if (field.Type != 0x20 || field.Length != 4 || previous <= 0 || value <= 0 || field.Value != previous
                || data.U32(field.Offset) != field.Value) return false;
            if (previous != value) edits.Add(new(field.Offset, 4, BitConverter.GetBytes((uint)value)));
            return true;
        }
        bool Axis(IReadOnlyList<double> oldTracks, IReadOnlyList<double> newTracks, int start, int oldExtent, int newExtent) {
            double oldEnd = 0, newEnd = 0; int oldEncodedEnd = 0, newEncodedEnd = 0;
            for (int i = 0; i < oldTracks.Count; i++) {
                context.Record(); PublisherBlock track = tracks[start + i];
                if (track.Id != 0 || track.Type != 0x88) return false;
                IReadOnlyList<PublisherBlock> fields = reader.Children(track);
                PublisherBlock[] ends = fields.Where(field => field.Id == 1).ToArray(), sizes = fields.Where(field => field.Id == 2).ToArray();
                if (fields.Count != 2 || ends.Length != 1 || sizes.Length != 1) return false;
                oldEnd += oldTracks[i]; newEnd += newTracks[i];
                if (!Emu(oldEnd, 0, out oldEncodedEnd) || !Emu(newEnd, 0, out newEncodedEnd)
                    || !Emu(oldTracks[i], 0, out int oldSize) || !Emu(newTracks[i], 0, out int newSize)
                    || !Scalar(ends[0], oldEncodedEnd, newEncodedEnd) || !Scalar(sizes[0], oldSize, newSize)) return false;
            }
            return oldEncodedEnd == oldExtent && newEncodedEnd == newExtent;
        }
    }
}
