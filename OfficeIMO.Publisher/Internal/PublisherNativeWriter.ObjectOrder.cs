namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<uint, IReadOnlyList<uint>> _objectOrders = new();

    private static bool SameObjectSet(IReadOnlyList<uint> a, IReadOnlyList<uint> b) =>
        a.Count == b.Count && new HashSet<uint>(a).SetEquals(b);

    /// <summary>Rejects crossing a text frame with artwork whose native wrapping is not explicitly disabled.</summary>
    internal static bool HasUnqualifiedWrapCrossing(PublisherContents contents, PublisherPublicationState publication,
        PublisherPageState before, IReadOnlyList<uint> order, CancellationToken token) {
        var positions = new Dictionary<uint, int>();
        for (int index = 0; index < order.Count; index++) {
            token.ThrowIfCancellationRequested(); positions.Add(order[index], index);
        }
        // Prefix maxima and suffix minima detect a frame/artwork inversion in
        // linear time, including empty native lists when wrapping starts behind text.
        var preceding = new int[order.Count]; var following = new int[order.Count];
        int maximum = -1, minimum = order.Count;
        for (int index = 0; index < before.ObjectIds.Count; index++) {
            token.ThrowIfCancellationRequested(); preceding[index] = maximum;
            uint id = before.ObjectIds[index];
            if (publication.Text.Frames.ContainsKey(id)) maximum = Math.Max(maximum, positions[id]);
        }
        for (int index = before.ObjectIds.Count - 1; index >= 0; index--) {
            token.ThrowIfCancellationRequested(); following[index] = minimum;
            uint id = before.ObjectIds[index];
            if (publication.Text.Frames.ContainsKey(id)) minimum = Math.Min(minimum, positions[id]);
        }
        for (int index = 0; index < before.ObjectIds.Count; index++) {
            token.ThrowIfCancellationRequested(); uint id = before.ObjectIds[index];
            if (preceding[index] < positions[id] && following[index] > positions[id]) continue;
            if (!contents.Shapes.TryGetValue(id, out PublisherSourceShape? shape)) return true;
            PublisherBlock[] wrap = shape.Fields.Where(field => field.Id is 4 or 0x12).ToArray();
            // Independent None/Square controls distinguish the exact None token
            // from an enabled obstacle with no current frame-exclusion references.
            if (wrap.Length != 1 || wrap[0].Id != 4 || wrap[0].Type != 0x10 || wrap[0].Value != 0x100) return true;
        }
        return false;
    }

    private void CompareObjectOrders() {
        for (int index = 0; index < _baseline.Pages.Count; index++) {
            _token.ThrowIfCancellationRequested();
            PublisherPageState before = _baseline.Pages[index], after = _current.Pages[index];
            if (before.ObjectIds.SequenceEqual(after.ObjectIds)) continue;
            if (HasUnqualifiedWrapCrossing(_imported.Contents, _baseline, before, after.ObjectIds, _token)) {
                Reject("Contents/page/" + before.Id, "Object ordering across native wrapping artwork and text frames is not qualified.");
                return;
            }
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
