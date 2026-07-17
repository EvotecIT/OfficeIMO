namespace OfficeIMO.OneNote;

internal static partial class OneNotePackageStoreReader {
    private static PackageIndex ReadStorageIndex(Stream stream, PackageDataElement element, OneNoteReaderOptions options) {
        var index = new PackageIndex();
        foreach (FssHttpStreamObject child in element.Node.Children) {
            if (child.Type == 0x0E) {
                byte[] data = FssHttpStreamObjectReader.ReadData(stream, child, 256, "storage-index cell mapping");
                var cursor = new FssHttpDataCursor(data, child.DataOffset);
                FssHttpCellId cellId = cursor.ReadCellId();
                OneNoteExtendedGuid manifestId = cursor.ReadExtendedGuid();
                cursor.SkipSerialNumber();
                cursor.EnsureEnd("storage-index cell mapping");
                index.CellMappings.Add(new PackageCellMapping(cellId, manifestId));
            } else if (child.Type == 0x0D) {
                byte[] data = FssHttpStreamObjectReader.ReadData(stream, child, 256, "storage-index revision mapping");
                var cursor = new FssHttpDataCursor(data, child.DataOffset);
                OneNoteExtendedGuid revisionId = cursor.ReadExtendedGuid();
                OneNoteExtendedGuid manifestId = cursor.ReadExtendedGuid();
                cursor.SkipSerialNumber();
                cursor.EnsureEnd("storage-index revision mapping");
                PackageGuidKey key = new PackageGuidKey(revisionId);
                if (index.RevisionMappings.ContainsKey(key)) {
                    throw new OneNoteFormatException("ONENOTE_PACKAGE_REVISION_MAPPING", "The storage index contains a duplicated revision mapping.", child.DataOffset);
                }
                index.RevisionMappings.Add(key, manifestId);
            }
        }
        if (index.CellMappings.Count == 0 || index.RevisionMappings.Count == 0) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_STORAGE_INDEX", "The storage index does not contain the required cell and revision mappings.", element.Node.HeaderOffset);
        }
        return index;
    }

    private sealed class PackageIndex {
        internal List<PackageCellMapping> CellMappings { get; } = new List<PackageCellMapping>();
        internal Dictionary<PackageGuidKey, OneNoteExtendedGuid> RevisionMappings { get; } = new Dictionary<PackageGuidKey, OneNoteExtendedGuid>();
    }

    private sealed class PackageCellMapping {
        internal PackageCellMapping(FssHttpCellId cellId, OneNoteExtendedGuid cellManifestElementId) {
            CellId = cellId;
            CellManifestElementId = cellManifestElementId;
        }
        internal FssHttpCellId CellId { get; }
        internal OneNoteExtendedGuid CellManifestElementId { get; }
    }
}
