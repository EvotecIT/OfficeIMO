namespace OfficeIMO.OneNote;

internal static class OneNoteTocMapper {
    private const uint JcidPersistablePropertyContainerForToc = 0x00020001;
    private const uint TocEntryIndexOidIndex = 0x24001CF6;
    private const uint NotebookColor = 0x14001CBE;
    private const uint EnableHistory = 0x08001E1E;
    private const uint FileIdentityGuid = 0x1C001D94;
    private const uint NotebookElementOrderingId = 0x14001CB9;
    private const uint FolderChildFilename = 0x1C001D6B;

    internal static OneNoteTocData Map(OneNoteRevisionStore store) {
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        OneNoteMaterializedObjectSpace space = materializer.FindCurrentSpaceByRootJcid(
            JcidPersistablePropertyContainerForToc,
            "ONENOTE_TOC_OBJECT_SPACE",
            "No current table-of-contents object space could be materialized.");
        OneNoteRevisionStoreObject root = space.GetRoot(1)
            ?? throw new OneNoteFormatException("ONENOTE_TOC_ROOT", "The table-of-contents root object is missing.");

        var result = new OneNoteTocData {
            ColorArgb = OneNoteSemanticMapper.ReadUInt32(root, NotebookColor),
            HistoryEnabled = OneNoteSemanticMapper.ReadBoolean(root, EnableHistory),
            RootObjectId = root.Id,
            StorageFormat = store.Header.StorageFormat
        };
        foreach (OneNoteExtendedGuid entryId in OneNoteSemanticMapper.GetReferences(root, TocEntryIndexOidIndex)) {
            OneNoteRevisionStoreObject? item = space.GetObject(entryId);
            if (item?.Jcid.Value != JcidPersistablePropertyContainerForToc) continue;
            string? name = OneNoteSemanticMapper.ReadString(item, FolderChildFilename);
            if (string.IsNullOrWhiteSpace(name)) continue;
            result.Entries.Add(new OneNoteTocEntry {
                Id = ReadGuid(item, FileIdentityGuid),
                Name = name!,
                Order = OneNoteSemanticMapper.ReadUInt32(item, NotebookElementOrderingId) ?? uint.MaxValue,
                ColorArgb = NormalizeColor(OneNoteSemanticMapper.ReadUInt32(item, NotebookColor))
            });
        }
        foreach (OneNoteRevisionStoreObject item in space.Objects) {
            result.PreservedObjects.Add(OneNoteSemanticMapper.CreateOpaqueObject(item, result.PreservedObjects.Count));
        }
        return result;
    }

    private static Guid? ReadGuid(OneNoteRevisionStoreObject item, uint propertyId) {
        byte[]? data = OneNoteSemanticMapper.ReadData(item, propertyId);
        return data != null && data.Length == 16 ? new Guid(data) : (Guid?)null;
    }

    private static uint? NormalizeColor(uint? value) => value == 0xFFFFFFFFU ? null : value;
}

internal sealed class OneNoteTocData {
    internal uint? ColorArgb { get; set; }
    internal bool? HistoryEnabled { get; set; }
    internal OneNoteExtendedGuid? RootObjectId { get; set; }
    internal OneNoteStorageFormat StorageFormat { get; set; }
    internal List<OneNoteTocEntry> Entries { get; } = new List<OneNoteTocEntry>();
    internal List<OneNoteOpaqueObject> PreservedObjects { get; } = new List<OneNoteOpaqueObject>();
}

internal sealed class OneNoteTocEntry {
    internal Guid? Id { get; set; }
    internal string Name { get; set; } = string.Empty;
    internal uint Order { get; set; }
    internal uint? ColorArgb { get; set; }
    internal bool IsSection => string.Equals(Path.GetExtension(Name), ".one", StringComparison.OrdinalIgnoreCase);
}
