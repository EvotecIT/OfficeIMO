using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore {
    private static readonly MapiProperty[] HierarchyColumns = {
        Column(0x0E30, MapiPropertyType.Integer32), Column(0x0E33, MapiPropertyType.Integer64),
        Column(0x0E34, MapiPropertyType.Binary), Column(0x0E38, MapiPropertyType.Integer32),
        Column(0x3001, MapiPropertyType.Unicode), Column(0x3602, MapiPropertyType.Integer32),
        Column(0x3603, MapiPropertyType.Integer32), Column(0x360A, MapiPropertyType.Boolean),
        Column(0x3613, MapiPropertyType.Unicode), Column(0x6635, MapiPropertyType.Integer32),
        Column(0x6636, MapiPropertyType.Integer32)
    };
    private static readonly MapiProperty[] ContentsColumns = {
        Column(0x0017, MapiPropertyType.Integer32), Column(0x001A, MapiPropertyType.Unicode),
        Column(0x0036, MapiPropertyType.Integer32), Column(0x0037, MapiPropertyType.Unicode),
        Column(0x0039, MapiPropertyType.Time), Column(0x0042, MapiPropertyType.Unicode),
        Column(0x0057, MapiPropertyType.Boolean), Column(0x0058, MapiPropertyType.Boolean),
        Column(0x0070, MapiPropertyType.Unicode), Column(0x0071, MapiPropertyType.Binary),
        Column(0x0E03, MapiPropertyType.Unicode), Column(0x0E04, MapiPropertyType.Unicode),
        Column(0x0E06, MapiPropertyType.Time), Column(0x0E07, MapiPropertyType.Integer32),
        Column(0x0E08, MapiPropertyType.Integer32), Column(0x0E17, MapiPropertyType.Integer32),
        Column(0x0E30, MapiPropertyType.Integer32), Column(0x0E33, MapiPropertyType.Integer64),
        Column(0x0E34, MapiPropertyType.Binary), Column(0x0E38, MapiPropertyType.Integer32),
        Column(0x0E3C, MapiPropertyType.Binary), Column(0x0E3D, MapiPropertyType.Binary),
        Column(0x1097, MapiPropertyType.Integer32), Column(0x3008, MapiPropertyType.Time),
        Column(0x65C6, MapiPropertyType.Integer32)
    };
    private static readonly MapiProperty[] AssociatedColumns = {
        Column(0x001A, MapiPropertyType.Unicode), Column(0x0E07, MapiPropertyType.Integer32),
        Column(0x0E17, MapiPropertyType.Integer32), Column(0x3001, MapiPropertyType.Unicode),
        Column(0x6800, MapiPropertyType.Unicode), Column(0x6803, MapiPropertyType.Boolean),
        Column(0x6805, MapiPropertyType.MultipleInteger32), Column(0x7003, MapiPropertyType.Integer32),
        Column(0x7004, MapiPropertyType.Binary), Column(0x7005, MapiPropertyType.Binary)
    };
    private static readonly MapiProperty[] RecipientColumns = {
        Column(0x0C15, MapiPropertyType.Integer32), Column(0x0E0F, MapiPropertyType.Boolean),
        Column(0x0FF9, MapiPropertyType.Binary), Column(0x0FFE, MapiPropertyType.Integer32),
        Column(0x0FFF, MapiPropertyType.Binary), Column(0x3001, MapiPropertyType.Unicode),
        Column(0x3002, MapiPropertyType.Unicode), Column(0x3003, MapiPropertyType.Unicode),
        Column(0x300B, MapiPropertyType.Binary), Column(0x3900, MapiPropertyType.Integer32),
        Column(0x39FF, MapiPropertyType.String8), Column(0x3A40, MapiPropertyType.Boolean)
    };
    private static readonly MapiProperty[] AttachmentColumns = {
        Column(0x0E20, MapiPropertyType.Integer32), Column(0x3704, MapiPropertyType.Unicode),
        Column(0x3705, MapiPropertyType.Integer32), Column(0x370B, MapiPropertyType.Integer32)
    };

    private void WriteStoreStructure(CancellationToken cancellationToken,
        PstWriterItemJournal.PstWriterItemSortedReader items) {
        WriteTemplateTable(0x60D, HierarchyColumns, "template/hierarchy");
        WriteTemplateTable(0x60E, ContentsColumns, "template/contents");
        WriteTemplateTable(0x60F, AssociatedColumns, "template/associated");
        WriteTemplateTable(0x610, ContentsColumns, "template/search");
        WriteTemplateTable(0x692, RecipientColumns, "template/recipients");
        WriteTemplateTable(0x671, AttachmentColumns, "template/attachments");

        Dictionary<uint, FolderState[]> childrenByParent = _folders.Values
            .Where(item => item.ParentNid != item.Nid)
            .GroupBy(item => item.ParentNid)
            .ToDictionary(group => group.Key, group => group.OrderBy(item => item.Nid).ToArray());
        bool hasInbox = _folders.Values.Any(item =>
            item.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox);
        foreach (FolderState folder in _folders.Values.OrderBy(item => item.Nid)) {
            cancellationToken.ThrowIfCancellationRequested();
            WriteFolder(folder, items, childrenByParent, hasInbox);
        }

        PstWriterContextResult nameMap = PstPropertyContextWriter.Write(_file,
            _namedProperties.BuildProperties(), 65001, null, null, null,
            Report, "named-properties");
        _nodes.Add(new PstWriterNode(0x61, 0, nameMap.DataBid, nameMap.SubnodeBid));

        byte[] uid = _providerUid.ToByteArray();
        var storeProperties = new List<MapiProperty> {
            new MapiProperty(0x0FF9, MapiPropertyType.Binary, uid),
            new MapiProperty(0x3001, MapiPropertyType.Unicode, _options.DisplayName),
            new MapiProperty(0x35DF, MapiPropertyType.Integer32, 0x89),
            new MapiProperty(0x67FF, MapiPropertyType.Integer32, 0)
        };
        AddSpecialFolderEntryId(storeProperties, 0x35E0, EmailStoreSpecialFolderKind.IpmSubtree);
        AddSpecialFolderEntryId(storeProperties, 0x35E1, EmailStoreSpecialFolderKind.Inbox);
        AddSpecialFolderEntryId(storeProperties, 0x35E2, EmailStoreSpecialFolderKind.Outbox);
        AddSpecialFolderEntryId(storeProperties, 0x35E3, EmailStoreSpecialFolderKind.DeletedItems);
        AddSpecialFolderEntryId(storeProperties, 0x35E4, EmailStoreSpecialFolderKind.SentItems);
        AddSpecialFolderEntryId(storeProperties, 0x35E5, EmailStoreSpecialFolderKind.PersonalViews);
        AddSpecialFolderEntryId(storeProperties, 0x35E6, EmailStoreSpecialFolderKind.CommonViews);
        AddSpecialFolderEntryId(storeProperties, 0x35E7, EmailStoreSpecialFolderKind.SearchRoot);
        PstWriterContextResult store = PstPropertyContextWriter.Write(_file,
            storeProperties, 65001, null, null, null, Report, "store");
        _nodes.Add(new PstWriterNode(0x21, 0, store.DataBid, store.SubnodeBid));

        ulong emptyQueue = _file.WriteDataTree(Array.Empty<byte>());
        _nodes.Add(new PstWriterNode(0x1E1, 0, emptyQueue));
        _nodes.Add(new PstWriterNode(0x201, 0, emptyQueue));
    }

    private void WriteFolder(FolderState folder,
        PstWriterItemJournal.PstWriterItemSortedReader items,
        IReadOnlyDictionary<uint, FolderState[]> childrenByParent,
        bool hasInbox) {
        FolderState[] children = childrenByParent.TryGetValue(folder.Nid, out FolderState[]? nested)
            ? nested
            : Array.Empty<FolderState>();
        var folderProperties = new List<MapiProperty> {
            new MapiProperty(0x0FF9, MapiPropertyType.Binary, CreateEntryId(folder.Nid)),
            new MapiProperty(0x3001, MapiPropertyType.Unicode, folder.Name),
            new MapiProperty(0x3602, MapiPropertyType.Integer32, folder.NormalItemCount),
            new MapiProperty(0x3603, MapiPropertyType.Integer32, folder.UnreadItemCount),
            new MapiProperty(0x360A, MapiPropertyType.Boolean, children.Length > 0),
            new MapiProperty(0x3617, MapiPropertyType.Integer32, folder.AssociatedItemCount)
        };
        if (!string.IsNullOrWhiteSpace(folder.ContainerClass)) {
            folderProperties.Add(new MapiProperty(0x3613, MapiPropertyType.Unicode, folder.ContainerClass));
        }
        if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox ||
            (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root && !hasInbox)) {
            AddSpecialFolderEntryId(folderProperties, 0x36D0, EmailStoreSpecialFolderKind.Calendar);
            AddSpecialFolderEntryId(folderProperties, 0x36D1, EmailStoreSpecialFolderKind.Contacts);
            AddSpecialFolderEntryId(folderProperties, 0x36D2, EmailStoreSpecialFolderKind.Journal);
            AddSpecialFolderEntryId(folderProperties, 0x36D3, EmailStoreSpecialFolderKind.Notes);
            AddSpecialFolderEntryId(folderProperties, 0x36D4, EmailStoreSpecialFolderKind.Tasks);
            AddSpecialFolderEntryId(folderProperties, 0x36D7, EmailStoreSpecialFolderKind.Drafts);
        }
        PstWriterContextResult pc = PstPropertyContextWriter.Write(_file,
            folderProperties, 65001, null, null, null, Report,
            string.Concat("folder/", FormatId(folder.Nid)));
        _nodes.Add(new PstWriterNode(folder.Nid, folder.ParentNid, pc.DataBid, pc.SubnodeBid));

        if (folder.IsSearchFolder) {
            _nodes.Add(new PstWriterNode((folder.Nid & ~0x1FU) | 0x06, 0, 0));
            PstWriterContextResult criteria = PstPropertyContextWriter.Write(_file,
                new[] { new MapiProperty(0x660B, MapiPropertyType.Integer32, 0) },
                65001, null, null, null, Report,
                string.Concat("folder/", FormatId(folder.Nid), "/search-criteria"));
            _nodes.Add(new PstWriterNode((folder.Nid & ~0x1FU) | 0x07, 0,
                criteria.DataBid, criteria.SubnodeBid));
            WriteFolderTable((folder.Nid & ~0x1FU) | 0x10, folder.Nid,
                Array.Empty<PstWriterTableRow>(), ContentsColumns, "search-contents");
            return;
        }

        var hierarchyRows = children.Select(child => new PstWriterTableRow(child.Nid,
            new[] {
                new MapiProperty(0x3001, MapiPropertyType.Unicode, child.Name),
                new MapiProperty(0x3602, MapiPropertyType.Integer32, child.NormalItemCount),
                new MapiProperty(0x3603, MapiPropertyType.Integer32, child.UnreadItemCount),
                new MapiProperty(0x360A, MapiPropertyType.Boolean,
                    childrenByParent.ContainsKey(child.Nid)),
                new MapiProperty(0x3613, MapiPropertyType.Unicode, child.ContainerClass)
            })).ToArray();
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0D, folder.Nid,
            hierarchyRows, HierarchyColumns, "hierarchy");
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0E, folder.Nid,
            items.ReadRows(folder.Nid, associated: false),
            ContentsColumns, "contents");
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0F, folder.Nid,
            items.ReadRows(folder.Nid, associated: true),
            AssociatedColumns, "associated");
    }

    private void WriteFolderTable(uint nid, uint parentNid,
        IEnumerable<PstWriterTableRow> rows, IReadOnlyList<MapiProperty> columns, string kind) {
        PstWriterContextResult table = PstTableContextWriter.Write(_file, rows, 65001,
            columns, Report, string.Concat("folder/", FormatId(parentNid), "/", kind));
        // The owning folder relationship is represented by the folder PC and
        // hierarchy table. Outlook writes top-level table-context NBT entries
        // with a zero parent NID.
        _nodes.Add(new PstWriterNode(nid, 0, table.DataBid, table.SubnodeBid));
    }

    private void WriteTemplateTable(uint nid, IReadOnlyList<MapiProperty> columns, string location) {
        PstWriterContextResult table = PstTableContextWriter.Write(_file,
            Array.Empty<PstWriterTableRow>(), 65001, columns, Report, location);
        _nodes.Add(new PstWriterNode(nid, 0, table.DataBid, table.SubnodeBid));
    }

    private byte[] CreateEntryId(uint nid) {
        var bytes = new byte[24];
        Buffer.BlockCopy(_providerUid.ToByteArray(), 0, bytes, 4, 16);
        PstBinary.WriteUInt32(bytes, 20, nid);
        return bytes;
    }

    private static IReadOnlyList<MapiProperty> SelectTableProperties(
        IEnumerable<MapiProperty> properties, IReadOnlyList<MapiProperty> columns) {
        var ids = new HashSet<ushort>(columns.Select(item => item.PropertyId));
        return properties.Where(item => ids.Contains(item.PropertyId)).ToArray();
    }

    private static bool IsUnread(IEnumerable<MapiProperty> properties) {
        MapiProperty? flags = properties.LastOrDefault(item => item.PropertyId == 0x0E07);
        return !(flags?.Value is int value) || (value & 1) == 0;
    }

    private static MapiProperty Column(ushort id, MapiPropertyType type) =>
        new MapiProperty(id, type, null);

    private void AddSpecialFolderEntryId(ICollection<MapiProperty> properties, ushort propertyId,
        EmailStoreSpecialFolderKind kind) {
        FolderState? folder = _folders.Values.FirstOrDefault(item => item.SpecialFolderKind == kind);
        if (folder != null) {
            properties.Add(new MapiProperty(propertyId, MapiPropertyType.Binary,
                CreateEntryId(folder.Nid)));
        }
    }
}
