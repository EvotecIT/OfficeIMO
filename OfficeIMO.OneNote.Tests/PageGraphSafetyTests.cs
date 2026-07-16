namespace OfficeIMO.OneNote.Tests;

public sealed class PageGraphSafetyTests {
    [Fact]
    public void SharedConflictPageIsRetainedOnlyOnFirstDagBranchAndCanBeRewritten() {
        var section = new OneNoteSection { Name = "Shared conflicts" };
        var current = new OneNotePage { Title = "Current" };
        var first = new OneNotePage { Title = "First", IsConflictPage = true };
        var second = new OneNotePage { Title = "Second", IsConflictPage = true };
        var shared = new OneNotePage { Title = "Shared", IsConflictPage = true };
        first.ConflictPages.Add(shared);
        current.ConflictPages.Add(first);
        current.ConflictPages.Add(second);
        section.Pages.Add(current);

        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        AddConflictReferences(graph, second, shared.Id!);

        OneNoteSection loaded = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)),
            new OneNoteReaderOptions { PreserveUnknownData = false });

        OneNotePage loadedCurrent = Assert.Single(loaded.Pages);
        OneNotePage firstShared = Assert.Single(loadedCurrent.ConflictPages[0].ConflictPages);
        Assert.Equal("Shared", firstShared.Title);
        Assert.Empty(loadedCurrent.ConflictPages[1].ConflictPages);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded);
        OneNotePage reopened = Assert.Single(OneNoteSectionReader.Read(new MemoryStream(rewritten)).Pages);
        Assert.Equal("Shared", Assert.Single(reopened.ConflictPages[0].ConflictPages).Title);
        Assert.Empty(reopened.ConflictPages[1].ConflictPages);
    }

    [Fact]
    public void CyclicConflictReferenceIsPrunedAndCanBeRewritten() {
        var section = new OneNoteSection { Name = "Cyclic conflicts" };
        var current = new OneNotePage { Title = "Current" };
        var conflict = new OneNotePage { Title = "Conflict", IsConflictPage = true };
        current.ConflictPages.Add(conflict);
        section.Pages.Add(current);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        AddConflictReferences(graph, conflict, current.Id!);

        OneNotePage loaded = Assert.Single(OneNoteSectionReader.Read(
            new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)),
            new OneNoteReaderOptions { PreserveUnknownData = false }).Pages);

        Assert.Empty(Assert.Single(loaded.ConflictPages).ConflictPages);
        OneNotePage reopened = Assert.Single(OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(new OneNoteSection {
                Name = "Rewritten",
                Pages = { loaded }
            }))).Pages);
        Assert.Empty(Assert.Single(reopened.ConflictPages).ConflictPages);
    }

    [Fact]
    public void DuplicateVersionContextReferenceIsPrunedAndCanBeRewritten() {
        var section = new OneNoteSection { Name = "Duplicate history" };
        var current = new OneNotePage { Title = "Current" };
        current.VersionHistory.Add(new OneNotePage { Title = "Earlier", IsVersionHistoryPage = true });
        section.Pages.Add(current);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        DuplicateVersionContextReference(graph);

        OneNotePage loaded = Assert.Single(OneNoteSectionReader.Read(
            new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)),
            new OneNoteReaderOptions { PreserveUnknownData = false }).Pages);

        Assert.Equal("Earlier", Assert.Single(loaded.VersionHistory).Title);
        var rewrittenSection = new OneNoteSection { Name = "Rewritten history" };
        rewrittenSection.Pages.Add(loaded);
        OneNotePage reopened = Assert.Single(OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(rewrittenSection))).Pages);
        Assert.Equal("Earlier", Assert.Single(reopened.VersionHistory).Title);
    }

    [Fact]
    public void RelatedPageDepthLimitFailsBeforeUnboundedRecursion() {
        byte[] data = WriteConflictChain(4);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionReader.Read(new MemoryStream(data), new OneNoteReaderOptions {
                MaxPageRelationshipDepth = 3
            }));

        Assert.Equal("ONENOTE_PAGE_GRAPH_DEPTH", exception.Code);
    }

    [Fact]
    public void DistinctPageGraphNodeLimitBoundsMaterialization() {
        byte[] data = WriteConflictChain(3);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionReader.Read(new MemoryStream(data), new OneNoteReaderOptions {
                MaxPageGraphNodes = 2
            }));

        Assert.Equal("ONENOTE_PAGE_GRAPH_LIMIT", exception.Code);
    }

    [Fact]
    public void PageGraphLimitsAllowTheExactConfiguredBoundary() {
        byte[] data = WriteConflictChain(3);

        OneNotePage loaded = Assert.Single(OneNoteSectionReader.Read(
            new MemoryStream(data),
            new OneNoteReaderOptions {
                MaxPageGraphNodes = 3,
                MaxPageRelationshipDepth = 3
            }).Pages);

        Assert.Equal("Page 2", Assert.Single(Assert.Single(loaded.ConflictPages).ConflictPages).Title);
    }

    [Fact]
    public void InvalidRelatedSpacesConsumePageGraphBudgetBeforeMaterialization() {
        var section = new OneNoteSection { Name = "Invalid references" };
        var current = new OneNotePage { Title = "Current" };
        section.Pages.Add(current);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        OneNoteWriteObjectSpace extraSpace = AddNonPageObjectSpace(graph);
        AddConflictReferences(graph, current, graph.ObjectSpaces[0].Id, extraSpace.Id);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionReader.Read(
                new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)),
                new OneNoteReaderOptions { MaxPageGraphNodes = 2 }));

        Assert.Equal("ONENOTE_PAGE_GRAPH_LIMIT", exception.Code);
    }

    [Fact]
    public void InvalidVersionHistoryContextsConsumePageGraphBudgetBeforeMaterialization() {
        var section = new OneNoteSection { Name = "Invalid history" };
        var current = new OneNotePage { Title = "Current" };
        current.VersionHistory.Add(new OneNotePage { Title = "Earlier", IsVersionHistoryPage = true });
        section.Pages.Add(current);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        ReplaceVersionHistoryContexts(
            graph,
            current,
            new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17),
            new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionReader.Read(
                new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)),
                new OneNoteReaderOptions { MaxPageGraphNodes = 2 }));

        Assert.Equal("ONENOTE_PAGE_GRAPH_LIMIT", exception.Code);
    }

    private static byte[] WriteConflictChain(int pageCount) {
        var section = new OneNoteSection { Name = "Deep conflicts" };
        var root = new OneNotePage { Title = "Page 0" };
        OneNotePage parent = root;
        for (int index = 1; index < pageCount; index++) {
            var child = new OneNotePage { Title = "Page " + index, IsConflictPage = true };
            parent.ConflictPages.Add(child);
            parent = child;
        }
        section.Pages.Add(root);
        return OneNoteSectionWriter.Write(section);
    }

    private static OneNoteWriteObjectSpace AddNonPageObjectSpace(OneNoteWriteGraph graph) {
        var spaceId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var rootId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var space = new OneNoteWriteObjectSpace(spaceId, new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17));
        space.Objects.Add(new OneNoteWriteObject(rootId, OneNoteSchema.JcidSectionMetadata));
        space.Roots[1] = rootId;
        graph.ObjectSpaces.Add(space);
        return space;
    }

    private static void DuplicateVersionContextReference(OneNoteWriteGraph graph) {
        OneNoteWriteObjectSpace historySpace = graph.ObjectSpaces.Single(space =>
            space.Objects.Any(item => item.Jcid == OneNoteSchema.JcidVersionProxy));
        int proxyIndex = historySpace.Objects.ToList().FindIndex(item => item.Jcid == OneNoteSchema.JcidVersionProxy);
        OneNoteWriteObject proxy = historySpace.Objects[proxyIndex];
        OneNoteWriteProperty context = proxy.Properties.Single(property =>
            (property.RawId & 0x03FFFFFFU) == (OneNoteSchema.VersionHistoryGraphSpaceContextNodes & 0x03FFFFFFU));
        OneNoteExtendedGuid contextId = Assert.Single(context.References);
        OneNoteWriteProperty[] properties = proxy.Properties
            .Select(property => ReferenceEquals(property, context)
                ? new OneNoteWriteProperty(
                    context.RawId,
                    references: new[] { contextId, contextId },
                    referenceKind: OneNoteWriteReferenceKind.Context,
                    preserveRawId: true)
                : property)
            .ToArray();
        historySpace.Objects[proxyIndex] = new OneNoteWriteObject(proxy.Id, proxy.Jcid, properties);
    }

    private static void ReplaceVersionHistoryContexts(
        OneNoteWriteGraph graph,
        OneNotePage page,
        params OneNoteExtendedGuid[] contextIds) {
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces.Single(space =>
            space.Id.Equals(page.Id) && space.ContextId == null);
        int manifestIndex = pageSpace.Objects.ToList().FindIndex(item => item.Jcid == OneNoteSchema.JcidPageManifestNode);
        OneNoteWriteObject manifest = pageSpace.Objects[manifestIndex];
        OneNoteWriteProperty[] properties = manifest.Properties
            .Select(property =>
                (property.RawId & 0x03FFFFFFU) == (OneNoteSchema.VersionHistoryGraphSpaceContextNodes & 0x03FFFFFFU)
                    ? new OneNoteWriteProperty(
                        property.RawId,
                        references: contextIds,
                        referenceKind: OneNoteWriteReferenceKind.Context,
                        preserveRawId: true)
                    : property)
            .ToArray();
        pageSpace.Objects[manifestIndex] = new OneNoteWriteObject(manifest.Id, manifest.Jcid, properties);
    }

    private static void AddConflictReferences(
        OneNoteWriteGraph graph,
        OneNotePage parent,
        params OneNoteExtendedGuid[] childIds) {
        OneNoteWriteObjectSpace parentSpace = graph.ObjectSpaces.Single(space => space.Id.Equals(parent.Id));
        int manifestIndex = parentSpace.Objects.ToList().FindIndex(item => item.Jcid == OneNoteSchema.JcidPageManifestNode);
        OneNoteWriteObject manifest = parentSpace.Objects[manifestIndex];
        OneNoteWriteProperty[] properties = manifest.Properties
            .Where(property => (property.RawId & 0x03FFFFFFU) != (OneNoteSchema.ChildGraphSpaceElementNodes & 0x03FFFFFFU))
            .Concat(new[] {
                new OneNoteWriteProperty(
                    OneNoteSchema.ChildGraphSpaceElementNodes,
                    references: childIds,
                    referenceKind: OneNoteWriteReferenceKind.ObjectSpace)
            })
            .ToArray();
        parentSpace.Objects[manifestIndex] = new OneNoteWriteObject(manifest.Id, manifest.Jcid, properties);
    }
}
