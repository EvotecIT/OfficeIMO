namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<uint, IReadOnlyList<uint>> _objectOrders = new();

    private static bool SameObjectSet(IReadOnlyList<uint> a, IReadOnlyList<uint> b) =>
        a.Count == b.Count && new HashSet<uint>(a).SetEquals(b);

    private void CompareObjectOrders() {
        for (int index = 0; index < _baseline.Pages.Count; index++) {
            _token.ThrowIfCancellationRequested();
            PublisherPageState before = _baseline.Pages[index], after = _current.Pages[index];
            if (before.ObjectIds.SequenceEqual(after.ObjectIds)) continue;
            foreach (uint id in before.ObjectIds) {
                _token.ThrowIfCancellationRequested();
                if (Binding(id) == null) return;
                PublisherEscherShape shape = _artwork[id];
                if (shape.IsGroup || shape.HasPublicationGroupParent
                    || _baseline.Text.Frames.TryGetValue(id, out PublisherFrameState? frame) && frame.WrapObjectIds.Count != 0) {
                    Reject("Contents/page/" + before.Id, "Object ordering with grouped artwork or native text-wrap exclusion lists is not qualified.");
                    return;
                }
            }
            _objectOrders.Add(before.Id, after.ObjectIds);
        }
    }

    private bool PrepareObjectOrder(Dictionary<string, byte[]> streams, PublisherParseContext context) {
        if (_objectOrders.Count == 0) return true;
        byte[] source = streams.TryGetValue("Escher/EscherStm", out byte[]? changed) ? changed : _imported.NativeSource!.EscherBytes;
        byte[]? reordered = PublisherEscherOrderWriter.Rewrite(source, _objectOrders, context, _maximumBytes);
        if (reordered == null) {
            Reject("Escher/EscherStm", "Page objects do not have an unambiguous complete sibling-record ordering profile.");
            return false;
        }
        streams["Escher/EscherStm"] = reordered;
        return true;
    }
}
