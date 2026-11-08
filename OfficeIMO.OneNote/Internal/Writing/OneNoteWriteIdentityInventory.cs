namespace OfficeIMO.OneNote;

/// <summary>Reserves identities retained later in a validated model before seeded allocation begins.</summary>
internal sealed class OneNoteWriteIdentityInventory {
    private readonly HashSet<Guid> _ids = new HashSet<Guid>();

    internal static IEnumerable<Guid> ForSection(OneNoteSection section) {
        var inventory = new OneNoteWriteIdentityInventory();
        inventory.Add(section.Id);
        inventory.Opaque(section.UnknownObjects);
        foreach (var revision in section.Revisions) { inventory.Add(revision.Id); inventory.Add(revision.BaseRevisionId); }
        inventory.Space(section.PreservationState?.SectionSpace);
        foreach (OneNotePage page in section.Pages) inventory.Page(page, section.PreservationState);
        return inventory._ids;
    }

    internal static IEnumerable<Guid> ForNotebook(OneNoteNotebook notebook) {
        var inventory = new OneNoteWriteIdentityInventory();
        inventory.Add(notebook.Id);
        inventory.NotebookChildren(notebook.Sections, notebook.SectionGroups);
        return inventory._ids;
    }

    internal static IEnumerable<Guid> ForToc(IEnumerable<OneNoteOpaqueObject> objects, OneNoteExtendedGuid? root) {
        var inventory = new OneNoteWriteIdentityInventory();
        inventory.Add(root); inventory.Opaque(objects);
        return inventory._ids;
    }

    private void NotebookChildren(IEnumerable<OneNoteSection> sections, IEnumerable<OneNoteSectionGroup> groups) {
        foreach (var section in sections) Add(section.Id);
        foreach (var group in groups) { Add(group.Id); NotebookChildren(group.Sections, group.SectionGroups); }
    }

    private void Page(OneNotePage page, OneNoteSectionPreservationState? preservation) {
        Add(page.Id); Add(page.RevisionContextId);
        var ids = page.PreservationIds;
        Add(ids.ManifestId); Add(ids.MetadataId); Add(ids.RevisionMetadataId); Add(ids.PageNodeId);
        Add(ids.TitleNodeId); Add(ids.TitleOutlineId); Add(ids.TitleElementId); Add(ids.TitleTextId); Add(ids.VersionProxyId);
        Add(preservation?.GetPageSeriesId(page)); Add(preservation?.GetCachedPageMetadataId(page));
        Space(preservation?.GetPageSpace(page));
        Opaque(page.UnknownObjects);
        foreach (var revision in page.Revisions) { Add(revision.Id); Add(revision.BaseRevisionId); }
        foreach (OneNoteElement element in OneNoteElementTraversal.Enumerate(page)) {
            Add(element.Id); Add(element.Author?.ObjectId); Properties(element.UnknownProperties);
            foreach (var tag in element.Tags) Add(tag.DefinitionId);
            if (element is OneNoteParagraph paragraph) {
                Add(paragraph.ContentObjectId); Add(paragraph.Style.ObjectId); Add(paragraph.List?.ObjectId);
                foreach (var run in paragraph.Runs) Add(run.StyleObjectId);
            }
            if (element is OneNoteTable table) foreach (var row in table.Rows) {
                Add(row.ObjectId);
                foreach (var cell in row.Cells) { Add(cell.ObjectId); Properties(cell.UnknownProperties); }
            }
            if (element is OneNoteBinaryElement binary) { Add(binary.PayloadObjectId); Add(binary.PayloadFileDataId); }
            if (element is OneNoteImage image) { Add(image.PictureContainerObjectId); Add(image.WebPictureContainerObjectId); }
            if (element is OneNoteInk ink) {
                Add(ink.InkDataObjectId);
                foreach (var id in ink.StrokeObjectIds.Values.Concat(ink.StrokePropertyObjectIds.Values).Concat(ink.PreservedStrokeObjectIds)) Add(id);
            }
            if (element is OneNoteMath math) Add(math.ContentObjectId);
            if (element is OneNoteMedia media) Add(media.RecordingId);
        }
        foreach (var conflict in page.ConflictPages) Page(conflict, preservation);
        foreach (var version in page.VersionHistory) Page(version, preservation);
    }

    private void Space(OneNoteMaterializedObjectSpace? space) {
        if (space == null) return;
        Add(space.Revision.Id); Add(space.Revision.ObjectSpaceId);
        foreach (var item in space.Objects) {
            Add(item.Id); Add(item.RevisionId);
            byte[]? management = OneNoteSemanticMapper.ReadData(item, OneNoteSchema.NotebookManagementEntityGuid);
            if (management?.Length == 16) Add(new Guid(management));
        }
    }

    private void Opaque(IEnumerable<OneNoteOpaqueObject> objects) {
        foreach (var item in objects) { Add(item.Id); Properties(item.Properties); }
    }
    private void Properties(IEnumerable<OneNoteOpaqueProperty> properties) {
        foreach (var property in properties) foreach (var id in property.ReferencedIds) Add(id);
    }
    private void Add(OneNoteExtendedGuid? id) { if (id != null) Add(id.Identifier); }
    private void Add(Guid? id) { if (id.HasValue && id.Value != Guid.Empty) _ids.Add(id.Value); }
}
