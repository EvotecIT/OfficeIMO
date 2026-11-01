namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal PublisherPublicationState SetPageObjectOrder(uint pageId, IReadOnlyList<uint> ids, CancellationToken token) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (ids.Count != page.ObjectIds.Count)
            throw new ArgumentException("Object order must contain every current page object exactly once.", nameof(ids));
        var remaining = new HashSet<uint>(page.ObjectIds);
        var order = new uint[ids.Count];
        for (int index = 0; index < order.Length; index++) {
            token.ThrowIfCancellationRequested();
            uint id = ids[index];
            if (!remaining.Remove(id))
                throw new ArgumentException("Object order contains a duplicate or an identifier outside this page.", nameof(ids));
            order[index] = id;
        }
        if (page.ObjectIds.SequenceEqual(order)) return this;
        return WithPages(Pages.Select(item => item.Id == pageId ? item.WithObjectOrder(order) : item), NextObjectId);
    }
}
