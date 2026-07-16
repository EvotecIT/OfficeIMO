using System.Text;

namespace OfficeIMO.OneNote;

internal sealed partial class OneNoteWriteGraphBuilder {
    private static readonly OneNoteExtendedGuid VersionHistoryContext = new OneNoteExtendedGuid(
        new Guid("7111497F-1B6B-4209-9491-C98B04CF4C5A"),
        1,
        17);
    private readonly OneNoteWriteIdFactory _ids = new OneNoteWriteIdFactory();
    private readonly long _maxPayloadBytes;
    private readonly bool _preserveUnknownData;

    internal OneNoteWriteGraphBuilder(
        long maxPayloadBytes = OneNoteReaderOptions.DefaultMaxInputBytes,
        bool preserveUnknownData = true) {
        _maxPayloadBytes = maxPayloadBytes;
        _preserveUnknownData = preserveUnknownData;
    }

    internal OneNoteWriteGraph BuildSection(OneNoteSection section, Guid? ancestorId = null, string? fileName = null, Guid? fileIdOverride = null) {
        OneNoteSectionPreservationState? preservation = _preserveUnknownData ? section.PreservationState : null;
        OneNoteMaterializedObjectSpace? sourceSectionSpace = preservation?.SectionSpace;
        Guid fileId = fileIdOverride ?? (section.Id.HasValue && section.Id.Value != Guid.Empty ? section.Id.Value : Guid.NewGuid());
        section.Id = fileId;
        OneNoteExtendedGuid sectionSpaceId = IdOrNew(sourceSectionSpace?.Revision.ObjectSpaceId);
        var graph = new OneNoteWriteGraph(fileId, OneNoteFileKind.Section, sectionSpaceId, ancestorId ?? Guid.Empty, OneNoteCrc32.ComputeFileName(fileName));
        var sectionSpace = new OneNoteWriteObjectSpace(sectionSpaceId, IdOrNew(sourceSectionSpace?.Revision.Id));
        var pageSeriesIds = new List<OneNoteExtendedGuid>();
        DateTime sectionCreationUtc = section.Pages
            .Where(page => page.CreatedUtc.HasValue)
            .Select(page => page.CreatedUtc!.Value.ToUniversalTime())
            .DefaultIfEmpty(DateTime.UtcNow)
            .Min();

        foreach (OneNotePage page in section.Pages) {
            OneNoteMaterializedObjectSpace? sourcePageSpace = preservation?.GetPageSpace(page);
            OneNoteExtendedGuid pageSpaceId = IdOrNew(sourcePageSpace?.Revision.ObjectSpaceId ?? page.Id);
            Guid pageManagementId = ReadGuidProperty(sourcePageSpace?.GetRoot(2), OneNoteSchema.NotebookManagementEntityGuid) ?? Guid.NewGuid();
            DateTime pageCreationUtc = page.CreatedUtc?.ToUniversalTime() ?? sectionCreationUtc;
            var conflictSpaceIds = new List<OneNoteExtendedGuid>();
            foreach (OneNotePage conflict in page.ConflictPages) {
                conflictSpaceIds.Add(BuildConflictPageSpaces(graph, conflict, sectionCreationUtc, preservation));
            }
            IReadOnlyList<OneNoteExtendedGuid> versionHistoryContextIds = BuildVersionHistorySpaces(
                graph,
                page,
                pageSpaceId,
                pageManagementId,
                pageCreationUtc,
                preservation);
            graph.ObjectSpaces.Add(BuildPageSpace(
                page,
                pageSpaceId,
                pageManagementId,
                pageCreationUtc,
                sourcePageSpace,
                preservation,
                conflictSpaceIds,
                versionHistoryContextIds));

            OneNoteExtendedGuid cachedMetadataId = IdOrNew(preservation?.GetCachedPageMetadataId(page));
            sectionSpace.Objects.Add(new OneNoteWriteObject(
                cachedMetadataId,
                OneNoteSchema.JcidPageMetadata,
                PageMetadataProperties(page, pageManagementId, pageCreationUtc)));

            OneNoteExtendedGuid seriesId = IdOrNew(preservation?.GetPageSeriesId(page));
            OneNoteRevisionStoreObject? sourceSeries = sourceSectionSpace?.GetObject(seriesId);
            sectionSpace.Objects.Add(new OneNoteWriteObject(seriesId, OneNoteSchema.JcidPageSeriesNode, new[] {
                Data(OneNoteSchema.NotebookManagementEntityGuid, (ReadGuidProperty(sourceSeries, OneNoteSchema.NotebookManagementEntityGuid) ?? Guid.NewGuid()).ToByteArray()),
                ObjectSpaceReferences(OneNoteSchema.ChildGraphSpaceElementNodes, pageSpaceId),
                Scalar(OneNoteSchema.TopologyCreationTimestamp, FileTime(pageCreationUtc)),
                ObjectReferences(OneNoteSchema.MetaDataObjectsAboveGraphSpace, cachedMetadataId)
            }));
            pageSeriesIds.Add(seriesId);
        }

        OneNoteRevisionStoreObject? sourceSectionRoot = sourceSectionSpace?.GetRoot(1);
        OneNoteExtendedGuid sectionRootId = IdOrNew(sourceSectionRoot?.Id);
        var sectionRootProperties = new List<OneNoteWriteProperty> {
            Data(OneNoteSchema.NotebookManagementEntityGuid, (ReadGuidProperty(sourceSectionRoot, OneNoteSchema.NotebookManagementEntityGuid) ?? Guid.NewGuid()).ToByteArray()),
            Scalar(OneNoteSchema.TopologyCreationTimestamp, FileTime(sectionCreationUtc))
        };
        if (pageSeriesIds.Count > 0) sectionRootProperties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, pageSeriesIds));
        sectionSpace.Objects.Add(new OneNoteWriteObject(sectionRootId, OneNoteSchema.JcidSectionNode, sectionRootProperties));
        OneNoteExtendedGuid metadataId = IdOrNew(sourceSectionSpace?.GetRoot(2)?.Id);
        var metadataProperties = new List<OneNoteWriteProperty> {
            Data(OneNoteSchema.SectionDisplayName, Unicode(section.Name)),
            Scalar(OneNoteSchema.NotebookColor, section.ColorArgb ?? 0xFFFFFFFFU),
            Scalar(OneNoteSchema.SchemaRevisionInOrderToRead, 40),
            Scalar(OneNoteSchema.SchemaRevisionInOrderToWrite, 40)
        };
        sectionSpace.Objects.Add(new OneNoteWriteObject(metadataId, OneNoteSchema.JcidSectionMetadata, metadataProperties));
        sectionSpace.Roots[1] = sectionRootId;
        sectionSpace.Roots[2] = metadataId;
        if (preservation != null && sourceSectionSpace != null) {
            OneNotePreservationWriter.MergeSpace(sectionSpace, sourceSectionSpace, preservation, _maxPayloadBytes);
        }
        graph.ObjectSpaces.Insert(0, sectionSpace);
        return graph;
    }

    internal OneNoteWriteGraph BuildTableOfContents(
        Guid fileId,
        Guid ancestorId,
        string fileName,
        IEnumerable<OneNoteTocWriteEntry> entries,
        uint? colorArgb,
        bool? historyEnabled,
        IEnumerable<OneNoteOpaqueObject>? preservedObjects = null,
        OneNoteExtendedGuid? preservedRootId = null) {
        OneNoteOpaqueObject[] preserved = _preserveUnknownData
            ? preservedObjects?.OrderBy(item => item.Ordinal).ToArray() ?? Array.Empty<OneNoteOpaqueObject>()
            : Array.Empty<OneNoteOpaqueObject>();
        OneNoteExtendedGuid spaceId = _ids.New();
        var graph = new OneNoteWriteGraph(fileId, OneNoteFileKind.TableOfContents, spaceId, ancestorId, OneNoteCrc32.ComputeFileName(fileName));
        var space = new OneNoteWriteObjectSpace(spaceId, _ids.New());
        var entryIds = new List<OneNoteExtendedGuid>();
        foreach (OneNoteTocWriteEntry entry in entries.OrderBy(item => item.Order)) {
            OneNoteOpaqueObject? retained = FindTableOfContentsEntry(preserved, entry.Id);
            OneNoteExtendedGuid id = retained?.Id ?? _ids.New();
            var properties = new List<OneNoteWriteProperty> {
                Data(OneNoteSchema.FileIdentityGuid, entry.Id.ToByteArray()),
                Scalar(OneNoteSchema.NotebookElementOrderingId, entry.Order),
                Data(OneNoteSchema.FolderChildFilename, Unicode(entry.Name)),
                Scalar(OneNoteSchema.NotebookColor, entry.ColorArgb ?? 0xFFFFFFFFU)
            };
            space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidPropertyContainer, properties));
            entryIds.Add(id);
        }
        OneNoteOpaqueObject? retainedRoot = FindTableOfContentsRoot(preserved, preservedRootId);
        OneNoteExtendedGuid rootId = retainedRoot?.Id ?? preservedRootId ?? _ids.New();
        var rootProperties = new List<OneNoteWriteProperty>();
        if (entryIds.Count > 0) rootProperties.Add(ObjectReferences(OneNoteSchema.TocEntryIndex, entryIds));
        if (colorArgb.HasValue) rootProperties.Add(Scalar(OneNoteSchema.NotebookColor, colorArgb.Value));
        if (historyEnabled.HasValue) rootProperties.Add(Boolean(OneNoteSchema.EnableHistory, historyEnabled.Value));
        space.Objects.Add(new OneNoteWriteObject(rootId, OneNoteSchema.JcidPropertyContainer, rootProperties));
        space.Roots[1] = rootId;
        if (preserved.Length > 0) OneNoteOpaquePreservationWriter.MergeSpace(space, preserved, _maxPayloadBytes);
        graph.ObjectSpaces.Add(space);
        return graph;
    }

    private OneNoteWriteObjectSpace BuildPageSpace(
        OneNotePage page,
        OneNoteExtendedGuid spaceId,
        Guid managementId,
        DateTime creationUtc,
        OneNoteMaterializedObjectSpace? sourceSpace,
        OneNoteSectionPreservationState? preservation,
        IReadOnlyList<OneNoteExtendedGuid>? conflictSpaceIds = null,
        IReadOnlyList<OneNoteExtendedGuid>? versionHistoryContextIds = null,
        bool forceConflict = false,
        OneNoteExtendedGuid? contextId = null) {
        page.Id = spaceId;
        var space = new OneNoteWriteObjectSpace(spaceId, IdOrNew(sourceSpace?.Revision.Id), contextId);
        uint lastModifiedTime = Time32(page.LastModifiedUtc?.ToUniversalTime() ?? creationUtc);
        var pageContentIds = new List<OneNoteExtendedGuid>();
        foreach (OneNoteOutline outline in page.Outlines) pageContentIds.Add(BuildOutline(space, outline, lastModifiedTime));
        foreach (OneNoteElement element in page.DirectContent) pageContentIds.Add(BuildOutlineChild(space, element, lastModifiedTime));
        OneNoteExtendedGuid titleId = BuildTitle(space, page, lastModifiedTime, Time32(creationUtc));

        OneNoteExtendedGuid pageNodeId = IdOrNew(page.PreservationIds.PageNodeId);
        page.PreservationIds.PageNodeId = pageNodeId;
        var pageProperties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            Data(OneNoteSchema.CachedTitleStringFromPage, Unicode(page.Title)),
            ObjectReferences(OneNoteSchema.StructureElementChildNodes, titleId)
        };
        if (pageContentIds.Count > 0) pageProperties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, pageContentIds));
        if (!string.IsNullOrWhiteSpace(page.OriginalAuthor)) pageProperties.Add(Data(OneNoteSchema.Author, Unicode(page.OriginalAuthor!)));
        if (page.Width.HasValue) pageProperties.Add(Float(OneNoteSchema.PageWidth, page.Width.Value));
        if (page.Height.HasValue) pageProperties.Add(Float(OneNoteSchema.PageHeight, page.Height.Value));
        space.Objects.Add(new OneNoteWriteObject(pageNodeId, OneNoteSchema.JcidPageNode, pageProperties));

        OneNoteExtendedGuid manifestId = IdOrNew(page.PreservationIds.ManifestId);
        page.PreservationIds.ManifestId = manifestId;
        var manifestProperties = new List<OneNoteWriteProperty> {
            ObjectReferences(OneNoteSchema.ContentChildNodes, pageNodeId)
        };
        if (conflictSpaceIds != null && conflictSpaceIds.Count > 0) {
            manifestProperties.Add(ObjectSpaceReferences(OneNoteSchema.ChildGraphSpaceElementNodes, conflictSpaceIds.ToArray()));
        }
        if (versionHistoryContextIds != null && versionHistoryContextIds.Count > 0) {
            manifestProperties.Add(ContextReferences(OneNoteSchema.VersionHistoryGraphSpaceContextNodes, versionHistoryContextIds.ToArray()));
        }
        space.Objects.Add(new OneNoteWriteObject(manifestId, OneNoteSchema.JcidPageManifestNode, manifestProperties));
        OneNoteExtendedGuid metadataId = IdOrNew(page.PreservationIds.MetadataId);
        page.PreservationIds.MetadataId = metadataId;
        space.Objects.Add(new OneNoteWriteObject(
            metadataId,
            page.IsConflictPage || forceConflict ? OneNoteSchema.JcidConflictPageMetadata : OneNoteSchema.JcidPageMetadata,
            PageMetadataProperties(page, managementId, creationUtc)));
        OneNoteExtendedGuid revisionMetadataId = IdOrNew(page.PreservationIds.RevisionMetadataId);
        page.PreservationIds.RevisionMetadataId = revisionMetadataId;
        var revisionProperties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.LastModifiedTimestamp, FileTime(page.LastModifiedUtc?.ToUniversalTime() ?? creationUtc))
        };
        if (!string.IsNullOrWhiteSpace(page.MostRecentAuthor)) {
            revisionProperties.Add(ObjectReferences(OneNoteSchema.AuthorMostRecent, BuildAuthor(space, page.MostRecentAuthor!)));
        }
        space.Objects.Add(new OneNoteWriteObject(revisionMetadataId, OneNoteSchema.JcidRevisionMetadata, revisionProperties));
        space.Roots[1] = manifestId;
        space.Roots[2] = metadataId;
        space.Roots[4] = revisionMetadataId;
        if (preservation != null && sourceSpace != null) {
            OneNotePreservationWriter.MergeSpace(space, sourceSpace, preservation, _maxPayloadBytes);
        }
        return space;
    }

    private OneNoteExtendedGuid BuildConflictPageSpaces(
        OneNoteWriteGraph graph,
        OneNotePage page,
        DateTime sectionCreationUtc,
        OneNoteSectionPreservationState? preservation) {
        OneNoteMaterializedObjectSpace? sourceSpace = preservation?.GetPageSpace(page);
        OneNoteExtendedGuid spaceId = IdOrNew(sourceSpace?.Revision.ObjectSpaceId ?? page.Id);
        var childIds = new List<OneNoteExtendedGuid>();
        foreach (OneNotePage child in page.ConflictPages) {
            childIds.Add(BuildConflictPageSpaces(graph, child, sectionCreationUtc, preservation));
        }
        Guid managementId = ReadGuidProperty(sourceSpace?.GetRoot(2), OneNoteSchema.NotebookManagementEntityGuid) ?? Guid.NewGuid();
        DateTime creationUtc = page.CreatedUtc?.ToUniversalTime() ?? sectionCreationUtc;
        graph.ObjectSpaces.Add(BuildPageSpace(
            page,
            spaceId,
            managementId,
            creationUtc,
            sourceSpace,
            preservation,
            childIds,
            forceConflict: true));
        return spaceId;
    }

    private IReadOnlyList<OneNoteExtendedGuid> BuildVersionHistorySpaces(
        OneNoteWriteGraph graph,
        OneNotePage page,
        OneNoteExtendedGuid pageSpaceId,
        Guid managementId,
        DateTime pageCreationUtc,
        OneNoteSectionPreservationState? preservation) {
        if (page.VersionHistory.Count == 0) return Array.Empty<OneNoteExtendedGuid>();

        var versionContextIds = new List<OneNoteExtendedGuid>();
        foreach (OneNotePage version in page.VersionHistory) {
            OneNoteExtendedGuid versionContextId = version.RevisionContextId ?? _ids.New();
            if (versionContextId.Equals(VersionHistoryContext)) versionContextId = _ids.New();
            version.RevisionContextId = versionContextId;
            OneNoteMaterializedObjectSpace? sourceVersionSpace = preservation?.GetPageSpace(version);
            graph.ObjectSpaces.Add(BuildPageSpace(
                version,
                pageSpaceId,
                managementId,
                version.CreatedUtc?.ToUniversalTime() ?? pageCreationUtc,
                sourceVersionSpace,
                preservation,
                contextId: versionContextId));
            versionContextIds.Add(versionContextId);
        }

        OneNoteMaterializedObjectSpace? sourceHistorySpace = preservation?.Materializer.TryGetSpace(pageSpaceId, VersionHistoryContext);
        var historySpace = new OneNoteWriteObjectSpace(
            pageSpaceId,
            IdOrNew(sourceHistorySpace?.Revision.Id),
            VersionHistoryContext);
        var proxyIds = new List<OneNoteExtendedGuid>();
        for (int index = 0; index < page.VersionHistory.Count; index++) {
            OneNotePage version = page.VersionHistory[index];
            DateTime createdUtc = version.CreatedUtc?.ToUniversalTime() ?? pageCreationUtc;
            DateTime modifiedUtc = version.LastModifiedUtc?.ToUniversalTime() ?? createdUtc;
            var properties = new List<OneNoteWriteProperty> {
                Scalar(OneNoteSchema.CreationTimestamp, Time32(createdUtc)),
                Scalar(OneNoteSchema.LastModifiedTime, Time32(modifiedUtc)),
                Scalar(OneNoteSchema.LastModifiedTimestamp, FileTime(modifiedUtc)),
                ContextReferences(OneNoteSchema.VersionHistoryGraphSpaceContextNodes, versionContextIds[index])
            };
            if (!string.IsNullOrWhiteSpace(version.MostRecentAuthor)) {
                properties.Add(ObjectReferences(OneNoteSchema.AuthorMostRecent, BuildAuthor(historySpace, version.MostRecentAuthor!)));
            }
            OneNoteExtendedGuid proxyId = IdOrNew(version.PreservationIds.VersionProxyId);
            version.PreservationIds.VersionProxyId = proxyId;
            historySpace.Objects.Add(new OneNoteWriteObject(proxyId, OneNoteSchema.JcidVersionProxy, properties));
            proxyIds.Add(proxyId);
        }

        OneNoteExtendedGuid contentId = IdOrNew(sourceHistorySpace?.GetRoot(1)?.Id);
        historySpace.Objects.Add(new OneNoteWriteObject(
            contentId,
            OneNoteSchema.JcidVersionHistoryContent,
            proxyIds.Count == 0
                ? Array.Empty<OneNoteWriteProperty>()
                : new[] { ObjectReferences(OneNoteSchema.ElementChildNodes, proxyIds) }));
        OneNoteExtendedGuid metadataId = IdOrNew(sourceHistorySpace?.GetRoot(2)?.Id);
        historySpace.Objects.Add(new OneNoteWriteObject(metadataId, OneNoteSchema.JcidVersionHistoryMetadata, new[] {
            Scalar(OneNoteSchema.SchemaRevisionInOrderToRead, 40),
            Scalar(OneNoteSchema.SchemaRevisionInOrderToWrite, 40)
        }));
        historySpace.Roots[1] = contentId;
        historySpace.Roots[2] = metadataId;
        if (preservation != null && sourceHistorySpace != null) {
            OneNotePreservationWriter.MergeSpace(historySpace, sourceHistorySpace, preservation, _maxPayloadBytes);
        }
        graph.ObjectSpaces.Add(historySpace);
        return new[] { VersionHistoryContext };
    }

    private OneNoteExtendedGuid BuildTitle(
        OneNoteWriteObjectSpace space,
        OneNotePage page,
        uint lastModifiedTime,
        uint creationTime) {
        var richTextProperties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime)
        };
        if (page.Title.All(character => character <= 0x7F)) {
            richTextProperties.Add(Data(OneNoteSchema.TextExtendedAscii, Encoding.ASCII.GetBytes(page.Title)));
        } else {
            richTextProperties.Add(Data(OneNoteSchema.RichEditTextUnicode, Unicode(page.Title)));
        }
        richTextProperties.Add(Boolean(OneNoteSchema.IsTitleText, true));
        richTextProperties.Add(Scalar(OneNoteSchema.RichEditTextLanguageId, 0x0409));
        OneNoteExtendedGuid richTextId = IdOrNew(page.PreservationIds.TitleTextId);
        page.PreservationIds.TitleTextId = richTextId;
        space.Objects.Add(new OneNoteWriteObject(richTextId, OneNoteSchema.JcidRichTextNode, richTextProperties));

        OneNoteExtendedGuid elementId = IdOrNew(page.PreservationIds.TitleElementId);
        page.PreservationIds.TitleElementId = elementId;
        space.Objects.Add(new OneNoteWriteObject(elementId, OneNoteSchema.JcidOutlineElementNode, new[] {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            ObjectReferences(OneNoteSchema.ContentChildNodes, richTextId),
            Scalar(OneNoteSchema.OutlineElementChildLevel, 1),
            Scalar(OneNoteSchema.CreationTimestamp, creationTime),
            Boolean(OneNoteSchema.CannotBeSelected, true),
            Boolean(OneNoteSchema.IsTitleText, true)
        }));

        OneNoteExtendedGuid outlineId = IdOrNew(page.PreservationIds.TitleOutlineId);
        page.PreservationIds.TitleOutlineId = outlineId;
        space.Objects.Add(new OneNoteWriteObject(outlineId, OneNoteSchema.JcidOutlineNode, new[] {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            ObjectReferences(OneNoteSchema.ElementChildNodes, elementId),
            Data(OneNoteSchema.OutlineIndentDistances, OutlineIndentDistances()),
            Scalar(OneNoteSchema.BodyTextAlignment, 0),
            Boolean(OneNoteSchema.EnforceOutlineStructure, true),
            Scalar(OneNoteSchema.OutlineElementChildLevel, 1),
            Float(OneNoteSchema.LayoutMaxHeight, 0.6),
            Boolean(OneNoteSchema.CannotBeSelected, true),
            Boolean(OneNoteSchema.IsTitleText, true),
            Boolean(OneNoteSchema.DescendantsCannotBeMoved, true),
            Float(OneNoteSchema.LayoutMinimumOutlineWidth, 4.5),
            Boolean(OneNoteSchema.LayoutTightAlignment, true),
            Scalar(OneNoteSchema.LayoutAlignmentInParent, 0),
            Scalar(OneNoteSchema.LayoutAlignmentSelf, 12)
        }));

        OneNoteExtendedGuid titleId = IdOrNew(page.PreservationIds.TitleNodeId);
        page.PreservationIds.TitleNodeId = titleId;
        space.Objects.Add(new OneNoteWriteObject(titleId, OneNoteSchema.JcidTitleNode, new[] {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            ObjectReferences(OneNoteSchema.ElementChildNodes, outlineId),
            Float(OneNoteSchema.OffsetFromParentHorizontal, 0),
            Float(OneNoteSchema.OffsetFromParentVertical, 0),
            Scalar(OneNoteSchema.LayoutCollisionPriority, 0),
            Scalar(OneNoteSchema.LayoutAlignmentInParent, 0x0009000C),
            Scalar(OneNoteSchema.LayoutAlignmentSelf, 0)
        }));
        return titleId;
    }

    private static IReadOnlyList<OneNoteWriteProperty> PageMetadataProperties(
        OneNotePage page,
        Guid managementId,
        DateTime creationUtc) {
        var properties = new List<OneNoteWriteProperty> {
            Data(OneNoteSchema.CachedTitleString, Unicode(page.Title)),
            Data(OneNoteSchema.NotebookManagementEntityGuid, managementId.ToByteArray()),
            Scalar(OneNoteSchema.PageLevel, (uint)Math.Max(1, page.Level + 1)),
            Scalar(OneNoteSchema.SchemaRevisionInOrderToRead, 40),
            Scalar(OneNoteSchema.SchemaRevisionInOrderToWrite, 40),
            Scalar(OneNoteSchema.TopologyCreationTimestamp, FileTime(creationUtc))
        };
        if (page.IsDeleted) properties.Add(Data(OneNoteSchema.IsDeletedGraphSpaceContent, Array.Empty<byte>()));
        return properties;
    }
}
