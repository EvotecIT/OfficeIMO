namespace OfficeIMO.OneNote;

internal sealed partial class OneNoteWriteGraphBuilder {
    private static OneNoteOpaqueObject? FindTableOfContentsEntry(
        IEnumerable<OneNoteOpaqueObject> objects,
        Guid fileId) {
        foreach (OneNoteOpaqueObject item in objects) {
            OneNoteOpaqueProperty? identity = FindOpaqueProperty(item, OneNoteSchema.FileIdentityGuid);
            byte[]? data = identity?.GetRawData();
            if (data != null && data.Length == 16 && new Guid(data) == fileId) return item;
        }
        return null;
    }

    private static OneNoteOpaqueObject? FindTableOfContentsRoot(
        IEnumerable<OneNoteOpaqueObject> objects,
        OneNoteExtendedGuid? rootId) {
        if (rootId != null) {
            OneNoteOpaqueObject? exact = objects.FirstOrDefault(item => rootId.Equals(item.Id));
            if (exact != null) return exact;
        }

        return objects.FirstOrDefault(item => FindOpaqueProperty(item, OneNoteSchema.TocEntryIndex) != null);
    }

    private static OneNoteOpaqueProperty? FindOpaqueProperty(OneNoteOpaqueObject item, uint propertyId) {
        uint key = propertyId & 0x03FFFFFFU;
        return item.Properties.FirstOrDefault(property => (property.PropertyId & 0x03FFFFFFU) == key);
    }
}
