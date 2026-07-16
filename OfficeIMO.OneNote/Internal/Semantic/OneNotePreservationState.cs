namespace OfficeIMO.OneNote;

internal sealed class OneNoteSectionPreservationState {
    private readonly Dictionary<string, OneNoteMaterializedObjectSpace> _pageSpaces;
    private readonly Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid> _pageSeriesIds;
    private readonly Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid> _cachedPageMetadataIds;

    private OneNoteSectionPreservationState(
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteMaterializedObjectSpace sectionSpace,
        Dictionary<string, OneNoteMaterializedObjectSpace> pageSpaces,
        Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid> pageSeriesIds,
        Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid> cachedPageMetadataIds) {
        Materializer = materializer;
        SectionSpace = sectionSpace;
        _pageSpaces = pageSpaces;
        _pageSeriesIds = pageSeriesIds;
        _cachedPageMetadataIds = cachedPageMetadataIds;
        MappedObjectIds = new HashSet<OneNoteExtendedGuid>(materializer.MappedObjectIds);
    }

    internal OneNoteObjectSpaceMaterializer Materializer { get; }
    internal OneNoteMaterializedObjectSpace SectionSpace { get; }
    internal IReadOnlyCollection<OneNoteExtendedGuid> MappedObjectIds { get; }

    internal static OneNoteSectionPreservationState Capture(
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteMaterializedObjectSpace sectionSpace,
        IEnumerable<OneNotePage> pages) {
        var pageSpaces = new Dictionary<string, OneNoteMaterializedObjectSpace>(StringComparer.Ordinal);
        foreach (OneNotePage page in EnumeratePages(pages)) {
            if (page.Id == null) continue;
            OneNoteMaterializedObjectSpace? space = materializer.TryGetSpace(page.Id, page.RevisionContextId);
            if (space != null) pageSpaces[OneNoteObjectSpaceMaterializer.GetSpaceKey(page.Id, page.RevisionContextId)] = space;
        }

        var pageSeriesIds = new Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid>();
        var cachedMetadataIds = new Dictionary<OneNoteExtendedGuid, OneNoteExtendedGuid>();
        OneNoteRevisionStoreObject? sectionRoot = sectionSpace.GetRoot(1);
        foreach (OneNoteExtendedGuid seriesId in OneNoteSemanticMapper.GetReferences(sectionRoot, OneNoteSchema.ElementChildNodes)) {
            OneNoteRevisionStoreObject? series = sectionSpace.GetObject(seriesId);
            if (series?.Jcid.Value != OneNoteSchema.JcidPageSeriesNode) continue;
            OneNoteExtendedGuid[] pageSpaceIds = OneNoteSemanticMapper
                .GetReferences(series, OneNoteSchema.ChildGraphSpaceElementNodes)
                .ToArray();
            OneNoteExtendedGuid[] cachedIds = OneNoteSemanticMapper
                .GetReferences(series, OneNoteSchema.MetaDataObjectsAboveGraphSpace)
                .ToArray();
            for (int index = 0; index < pageSpaceIds.Length; index++) {
                OneNoteExtendedGuid pageSpaceId = pageSpaceIds[index];
                pageSeriesIds[pageSpaceId] = series.Id;
                if (index < cachedIds.Length) {
                    OneNoteExtendedGuid cachedId = cachedIds[index];
                    cachedMetadataIds[pageSpaceId] = cachedId;
                    sectionSpace.GetObject(cachedId);
                }
            }
        }
        sectionSpace.GetRoot(2);

        return new OneNoteSectionPreservationState(
            materializer,
            sectionSpace,
            pageSpaces,
            pageSeriesIds,
            cachedMetadataIds);
    }

    internal OneNoteMaterializedObjectSpace? GetPageSpace(OneNotePage page) {
        if (page.Id == null) return null;
        _pageSpaces.TryGetValue(OneNoteObjectSpaceMaterializer.GetSpaceKey(page.Id, page.RevisionContextId), out OneNoteMaterializedObjectSpace? space);
        return space;
    }

    internal OneNoteExtendedGuid? GetPageSeriesId(OneNotePage page) {
        if (page.Id == null) return null;
        _pageSeriesIds.TryGetValue(page.Id, out OneNoteExtendedGuid? id);
        return id;
    }

    internal OneNoteExtendedGuid? GetCachedPageMetadataId(OneNotePage page) {
        if (page.Id == null) return null;
        _cachedPageMetadataIds.TryGetValue(page.Id, out OneNoteExtendedGuid? id);
        return id;
    }

    private static IEnumerable<OneNotePage> EnumeratePages(IEnumerable<OneNotePage> pages) {
        foreach (OneNotePage page in pages) {
            yield return page;
            foreach (OneNotePage conflict in EnumeratePages(page.ConflictPages)) yield return conflict;
            foreach (OneNotePage version in EnumeratePages(page.VersionHistory)) yield return version;
        }
    }
}
