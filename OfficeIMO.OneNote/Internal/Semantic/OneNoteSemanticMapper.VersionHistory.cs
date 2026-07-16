namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private const uint JcidVersionHistoryContent = 0x0006003C;
    private const uint JcidVersionProxy = 0x0006003D;
    private const uint VersionHistoryGraphSpaceContextNodes = 0x3400347B;
    private const uint CreationTimestamp = 0x14001D09;

    private static void MapVersionHistory(
        OneNotePage page,
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteExtendedGuid objectSpaceId,
        OneNoteRevisionStoreObject manifest,
        OneNoteReaderOptions options,
        PageMappingState pageMapping) {
        foreach (OneNoteExtendedGuid historyContextId in GetReferences(manifest, VersionHistoryGraphSpaceContextNodes)) {
            string historyKey = OneNoteObjectSpaceMaterializer.GetSpaceKey(objectSpaceId, historyContextId);
            if (!pageMapping.TryVisit(historyKey)) continue;
            OneNoteMaterializedObjectSpace? historySpace = materializer.TryGetSpace(objectSpaceId, historyContextId);
            OneNoteRevisionStoreObject? historyRoot = historySpace?.GetRoot(1);
            if (historySpace == null || historyRoot?.Jcid.Value != JcidVersionHistoryContent) continue;

            foreach (OneNoteExtendedGuid proxyId in GetReferences(historyRoot, ElementChildNodes)) {
                OneNoteRevisionStoreObject? proxy = historySpace.GetObject(proxyId);
                if (proxy?.Jcid.Value != JcidVersionProxy) continue;

                foreach (OneNoteExtendedGuid versionContextId in GetReferences(proxy, VersionHistoryGraphSpaceContextNodes)) {
                    OneNotePage? version = MapPage(materializer, objectSpaceId, versionContextId, true, options, pageMapping);
                    if (version == null) continue;

                    version.PreservationIds.VersionProxyId = proxy.Id;
                    version.CreatedUtc = ReadTime32(proxy, CreationTimestamp) ?? version.CreatedUtc;
                    version.LastModifiedUtc = ReadFileTime(proxy, LastModifiedTimestamp)
                        ?? ReadTime32(proxy, LastModifiedTime)
                        ?? version.LastModifiedUtc;
                    version.MostRecentAuthor = ReadReferencedAuthor(historySpace, proxy, AuthorMostRecent)
                        ?? version.MostRecentAuthor;
                    page.VersionHistory.Add(version);
                }
            }
        }
    }
}
