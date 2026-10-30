using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<int, IReadOnlyDictionary<ushort, uint>> _frameProperties = new();
    private readonly Dictionary<uint, uint> _frameAlignments = new();

    private void FrameLayout(PublisherFrameState before, PublisherFrameState after) {
        if (before.ColumnCount == after.ColumnCount && before.ColumnSpacing == after.ColumnSpacing
            && before.Padding.Equals(after.Padding) && before.VerticalAlignment == after.VerticalAlignment) return;
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
        if (changes.Count > 0) _frameProperties.Add(binding.BodyOffset, changes);
        if (before.VerticalAlignment != after.VerticalAlignment) _frameAlignments.Add(before.Id, (uint)after.VerticalAlignment);
    }

    private bool PrepareFrameLayoutStreams(Dictionary<string, byte[]> streams, byte[] patchedEscher) {
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = _sourceLimits }, _token);
        if (_frameProperties.Count > 0) {
            byte[]? artwork = PublisherEscherLayoutWriter.Rewrite(patchedEscher, _frameProperties, context, _maximumBytes);
            if (artwork == null) { Reject("Escher/EscherStm", "Frame layout has ambiguous or unsupported local OfficeArt property tables."); return false; }
            streams["Escher/EscherStm"] = artwork;
        }
        if (_frameAlignments.Count > 0) {
            byte[]? contents = PublisherContentsLayoutWriter.Rewrite(_imported.NativeSource!.ContentsBytes,
                _imported.Contents, _frameAlignments, context, _maximumBytes);
            if (contents == null) { Reject("Contents", "Frame alignment has an unsupported shape or chunk-directory relocation profile."); return false; }
            streams.Add("Contents", contents);
        }
        return true;
    }
}
