using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>> _shapeProperties = new();
    private readonly Dictionary<uint, uint> _frameAlignments = new();
    private bool _hasFrameLayout;

    private void FrameLayout(PublisherFrameState before, PublisherFrameState after) {
        if (before.ColumnCount == after.ColumnCount && before.ColumnSpacing == after.ColumnSpacing
            && before.Padding.Equals(after.Padding) && before.VerticalAlignment == after.VerticalAlignment) return;
        _hasFrameLayout = true;
        bool resolved = _baseline.Text.Chains.TryGetValue(before.StoryId, out var chain) && chain.Contains(before.Id)
            && chain.TakeWhile(id => id != before.Id).All(id => _baseline.Text.Frames[id].Available);
        if (_baseline.Text.FrameLayoutEditLimitation(before.Id, resolved) != null) {
            Reject("Contents/object/" + before.Id, "The source frame has no qualified layout edit."); return;
        }
        PublisherNativeShapeBinding? binding = Binding(before.Id);
        if (binding == null) return;
        var changes = new Dictionary<ushort, uint>();
        void Distance(ushort id, double prior, double value) {
            if (prior == value) return;
            if (value < 0 || !Emu(value, 0, out int encoded) || encoded < 0) {
                Reject("Contents/object/" + before.Id, "Frame spacing or insets have no exact nonnegative signed-EMU representation."); return;
            }
            changes.Add(id, (uint)encoded);
        }
        if (before.ColumnCount != after.ColumnCount) changes.Add(0x8C, checked((uint)after.ColumnCount));
        Distance(0x8D, before.ColumnSpacing, after.ColumnSpacing);
        Distance(0x81, before.Padding.Left, after.Padding.Left); Distance(0x82, before.Padding.Top, after.Padding.Top);
        Distance(0x83, before.Padding.Right, after.Padding.Right); Distance(0x84, before.Padding.Bottom, after.Padding.Bottom);
        if (changes.Count > 0) _shapeProperties.Add(binding.BodyOffset,
            changes.ToDictionary(pair => pair.Key, pair => new OfficeArtProperty(0, pair.Key, pair.Value)));
        if (before.VerticalAlignment != after.VerticalAlignment) _frameAlignments.Add(before.Id, (uint)after.VerticalAlignment);
    }

    private bool PrepareStructuredStreams(Dictionary<string, byte[]> streams, byte[] patchedEscher) {
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = _sourceLimits }, _token);
        if (_shapeProperties.Count > 0 || _imageStore != null) {
            byte[]? artwork = PublisherEscherSourceWriter.Rewrite(patchedEscher, _shapeProperties, _imageStore, context, _maximumBytes);
            if (artwork == null) { Reject("Escher/EscherStm", "Source edits have ambiguous or unsupported local OfficeArt records."); return false; }
            streams["Escher/EscherStm"] = artwork;
        }
        if (_frameAlignments.Count > 0 || _pictureNames.Count > 0 || _masterAssignments.Count > 0) {
            byte[]? contents = PublisherContentsSourceWriter.Rewrite(_imported.NativeSource!.ContentsBytes,
                _imported.Contents, _frameAlignments, _pictureNames, _masterAssignments, context, _maximumBytes);
            if (contents == null) { Reject("Contents", "Source edits have an unsupported page relationship, shape, picture metadata or chunk-directory relocation profile."); return false; }
            streams.Add("Contents", contents);
        }
        return true;
    }
}
