namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    internal PublisherModelState SetPageObjectOrder(uint pageId, IReadOnlyList<uint> ids, CancellationToken token) {
        PublisherPublicationState publication = _publication.SetPageObjectOrder(pageId, ids, token);
        if (ReferenceEquals(publication, _publication)) return this;
        PublisherPageState page = publication.Pages.Single(item => item.Id == pageId);
        var selected = new HashSet<uint>(page.ObjectIds);
        if (_imported != null && _imported.Artwork.Shapes.Any(shape => selected.Contains(shape.Id)
            && (shape.IsGroup || shape.HasPublicationGroupParent)))
            throw new NotSupportedException("Ordering grouped native artwork is not qualified.");
        if (_imported != null && PublisherNativeWriter.HasUnqualifiedWrapCrossing(_imported.Contents, _publication,
            _publication.Pages.Single(item => item.Id == pageId), page.ObjectIds, token))
            throw new NotSupportedException("Ordering across native wrapping artwork and text frames is not qualified.");
        return WithPublication(publication);
    }
}
