namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    private PublisherPublicationState RefreshWraps(PublisherPublicationState publication, CancellationToken token) {
        if (_imported?.Baseline == null) return publication;
        PublisherParseContext? context = null;
        var beforePages = _publication.Pages.ToDictionary(page => page.Id);
        var sourcePages = _imported.Baseline.Pages.ToDictionary(page => page.Id);
        foreach (PublisherPageState page in publication.Pages) {
            token.ThrowIfCancellationRequested();
            beforePages.TryGetValue(page.Id, out PublisherPageState? before);
            sourcePages.TryGetValue(page.Id, out PublisherPageState? source);
            if (before == null || source == null || !PublisherWrapRelationships.GeometryChanged(_publication, before, publication, page)) continue;
            if (before.ObjectIds.Count != page.ObjectIds.Count || !new HashSet<uint>(before.ObjectIds).SetEquals(page.ObjectIds)) continue;
            context ??= _decoded.ForProjection(token);
            if (PublisherWrapRelationships.TryCompute(_imported.Baseline, source, context, out var original)
                && PublisherWrapRelationships.Matches(_imported.Baseline, original)
                && PublisherWrapRelationships.TryCompute(publication, page, context, out var current))
                publication = publication.WithText(publication.Text.WithFrameWraps(current));
            else if (!before.ObjectIds.SequenceEqual(page.ObjectIds)
                && (_publication.Text.Frames.Values.Any(frame => frame.PageId == page.Id && frame.WrapObjectIds.Count > 0)
                    || PublisherNativeWriter.HasUnqualifiedWrapCrossing(_imported.Contents, _publication, before, page.ObjectIds, token)))
                throw new NotSupportedException("Object ordering requires unqualified native wrap relationships.");
        }
        return publication;
    }
}
