using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore {
    private static readonly MapiProperty[] HierarchyColumns = {
        Column(MapiKnownProperties.PidTag.ReplItemid), Column(0x0E30, MapiPropertyType.Binary),
        Column(MapiKnownProperties.PidTag.ReplChangenum),
        Column(MapiKnownProperties.PidTag.ReplVersionHistory), Column(MapiKnownProperties.PidTag.ReplFlags),
        Column(MapiKnownProperties.PidTag.DisplayName), Column(MapiKnownProperties.PidTag.ContentCount),
        Column(MapiKnownProperties.PidTag.ContentUnreadCount), Column(MapiKnownProperties.PidTag.Subfolders),
        Column(MapiKnownProperties.PidTag.ContainerClass), Column(MapiKnownProperties.PidTag.PstHiddenCount),
        Column(MapiKnownProperties.PidTag.PstHiddenUnread)
    };
    private static readonly MapiProperty[] ContentsColumns = {
        Column(MapiKnownProperties.PidTag.Importance), Column(MapiKnownProperties.PidTag.MessageClass),
        Column(MapiKnownProperties.PidTag.Sensitivity), Column(MapiKnownProperties.PidTag.Subject),
        Column(MapiKnownProperties.PidTag.ClientSubmitTime), Column(MapiKnownProperties.PidTag.SentRepresentingName),
        Column(MapiKnownProperties.PidTag.MessageToMe), Column(MapiKnownProperties.PidTag.MessageCcMe),
        Column(MapiKnownProperties.PidTag.ConversationTopic), Column(MapiKnownProperties.PidTag.ConversationIndex),
        Column(MapiKnownProperties.PidTag.DisplayCc), Column(MapiKnownProperties.PidTag.DisplayTo),
        Column(MapiKnownProperties.PidTag.MessageDeliveryTime), Column(MapiKnownProperties.PidTag.MessageFlags),
        Column(MapiKnownProperties.PidTag.MessageSize), Column(MapiKnownProperties.PidTag.MessageStatus),
        Column(MapiKnownProperties.PidTag.ReplItemid), Column(0x0E30, MapiPropertyType.Binary),
        Column(MapiKnownProperties.PidTag.ReplChangenum),
        Column(MapiKnownProperties.PidTag.ReplVersionHistory), Column(MapiKnownProperties.PidTag.ReplFlags),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromVersionhistory),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromItemid),
        Column(MapiKnownProperties.PidTag.ItemTemporaryFlags),
        Column(MapiKnownProperties.PidTag.LastModificationTime),
        Column(MapiKnownProperties.PidTag.ConversationId),
        Column(MapiKnownProperties.PidTag.SecureSubmitFlags)
    };
    private static readonly MapiProperty[] AssociatedColumns = {
        Column(MapiKnownProperties.PidTag.MessageClass), Column(MapiKnownProperties.PidTag.MessageFlags),
        Column(MapiKnownProperties.PidTag.MessageStatus), Column(MapiKnownProperties.PidTag.DisplayName),
        Column(MapiKnownProperties.PidTag.OfflineAddressBookName),
        Column(MapiKnownProperties.PidTag.SendOutlookRecallReport),
        Column(MapiKnownProperties.PidTag.OfflineAddressBookTruncatedProperties),
        Column(0x682F, MapiPropertyType.Unicode),
        Column(MapiKnownProperties.PidTag.ViewDescriptorFlags), Column(MapiKnownProperties.PidTag.ViewDescriptorLinkTo),
        Column(MapiKnownProperties.PidTag.ViewDescriptorViewFolder),
        Column(MapiKnownProperties.PidTag.ViewDescriptorName),
        Column(MapiKnownProperties.PidTag.ViewDescriptorVersion)
    };
    private static readonly MapiProperty[] SearchContentsColumns = {
        Column(MapiKnownProperties.PidTag.Importance), Column(MapiKnownProperties.PidTag.MessageClass),
        Column(MapiKnownProperties.PidTag.Sensitivity), Column(MapiKnownProperties.PidTag.Subject),
        Column(MapiKnownProperties.PidTag.ClientSubmitTime), Column(MapiKnownProperties.PidTag.SentRepresentingName),
        Column(MapiKnownProperties.PidTag.MessageToMe), Column(MapiKnownProperties.PidTag.MessageCcMe),
        Column(MapiKnownProperties.PidTag.ConversationTopic), Column(MapiKnownProperties.PidTag.ConversationIndex),
        Column(MapiKnownProperties.PidTag.DisplayCc), Column(MapiKnownProperties.PidTag.DisplayTo),
        Column(0x0E05, MapiPropertyType.Unicode), Column(MapiKnownProperties.PidTag.MessageDeliveryTime),
        Column(MapiKnownProperties.PidTag.MessageFlags), Column(MapiKnownProperties.PidTag.MessageSize),
        Column(MapiKnownProperties.PidTag.MessageStatus), Column(0x0E2A, MapiPropertyType.Boolean),
        Column(MapiKnownProperties.PidTag.ReplItemid), Column(MapiKnownProperties.PidTag.ReplChangenum),
        Column(MapiKnownProperties.PidTag.ReplVersionHistory), Column(MapiKnownProperties.PidTag.ReplFlags),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromVersionhistory),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromItemid),
        Column(MapiKnownProperties.PidTag.ItemTemporaryFlags),
        Column(MapiKnownProperties.PidTag.LastModificationTime),
        Column(MapiKnownProperties.PidTag.SecureSubmitFlags),
        Column(0x67F1, MapiPropertyType.Integer32)
    };
    private static readonly MapiProperty[] RecipientColumns = {
        Column(MapiKnownProperties.PidTag.RecipientType), Column(MapiKnownProperties.PidTag.Responsibility),
        Column(MapiKnownProperties.PidTag.RecordKey), Column(MapiKnownProperties.PidTag.ObjectType),
        Column(MapiKnownProperties.PidTag.EntryId), Column(MapiKnownProperties.PidTag.DisplayName),
        Column(MapiKnownProperties.PidTag.AddressType), Column(MapiKnownProperties.PidTag.EmailAddress),
        Column(MapiKnownProperties.PidTag.SearchKey), Column(MapiKnownProperties.PidTag.DisplayType),
        Column(MapiKnownProperties.PidTag.DisplayNamePrintable, MapiPropertyType.String8),
        Column(MapiKnownProperties.PidTag.DisplayNamePrintable, MapiPropertyType.Unicode),
        Column(MapiKnownProperties.PidTag.SendRichInfo)
    };
    private static readonly MapiProperty[] AttachmentColumns = {
        Column(MapiKnownProperties.PidTag.AttachSize), Column(MapiKnownProperties.PidTag.AttachFilename),
        Column(MapiKnownProperties.PidTag.AttachMethod), Column(MapiKnownProperties.PidTag.RenderingPosition)
    };
    private static readonly MapiProperty[] ChangeHistoryColumns = {
        Column(MapiKnownProperties.PidTag.ReplChangenum), Column(0x0E37, MapiPropertyType.Binary),
        Column(MapiKnownProperties.PidTag.ReplFlags)
    };
    private static readonly MapiProperty[] ReplicationColumns = {
        Column(MapiKnownProperties.PidTag.MessageClass), Column(0x0E30, MapiPropertyType.Binary),
        Column(0x0E31, MapiPropertyType.Binary), Column(MapiKnownProperties.PidTag.ReplChangenum),
        Column(MapiKnownProperties.PidTag.ReplVersionHistory), Column(MapiKnownProperties.PidTag.ReplFlags),
        Column(0x0E3E, MapiPropertyType.Binary)
    };
    private static readonly MapiProperty[] ChangeTrackingColumns = {
        Column(MapiKnownProperties.PidTag.ReplChangenum), Column(MapiKnownProperties.PidTag.CreationTime)
    };
    private static readonly MapiProperty[] ReceiveFolderColumns = {
        Column(MapiKnownProperties.PidTag.MessageClass), Column(0x6605, MapiPropertyType.Integer32)
    };
    private static readonly MapiProperty[] OutgoingQueueColumns = {
        Column(MapiKnownProperties.PidTag.ClientSubmitTime),
        Column(0x0E10, MapiPropertyType.Integer32), Column(0x0E14, MapiPropertyType.Integer32)
    };

    private void WriteStoreStructure(CancellationToken cancellationToken,
        PstWriterItemJournal.PstWriterItemSortedReader items) {
        WriteTemplateTable(0x60D, HierarchyColumns, "template/hierarchy");
        WriteTemplateTable(0x60E, ContentsColumns, "template/contents");
        WriteTemplateTable(0x60F, AssociatedColumns, "template/associated");
        WriteTemplateTable(0x610, SearchContentsColumns, "template/search");
        WriteTemplateTable(0x692, RecipientColumns, "template/recipients");
        WriteTemplateTable(0x671, AttachmentColumns, "template/attachments");
        WriteTemplateTable(0x6B6, ChangeHistoryColumns, "template/change-history");
        WriteTemplateTable(0x6D7, ReplicationColumns, "template/replication");
        WriteTemplateTable(0x6F8, ChangeTrackingColumns, "template/change-tracking");

        Dictionary<uint, FolderState[]> childrenByParent = _folders.Values
            .Where(item => item.ParentNid != item.Nid)
            .GroupBy(item => item.ParentNid)
            .ToDictionary(group => group.Key, group => group.OrderBy(item => item.Nid).ToArray());
        bool hasInbox = _folders.Values.Any(item =>
            item.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox);
        WriteReceiveFolderTable(hasInbox);
        WriteOutgoingQueueTable();
        foreach (FolderState folder in _folders.Values.OrderBy(item => item.Nid)) {
            cancellationToken.ThrowIfCancellationRequested();
            WriteFolder(folder, items, childrenByParent, hasInbox);
        }

        IReadOnlyList<MapiProperty> writerProvenance = _namedProperties.Map(new[] {
            new MapiProperty(0, MapiKnownProperties.PidName.OfficeImoPstWriter.PreferredType,
                PstWriterProvenance.PropertyValue, name: MapiKnownProperties.PidName.OfficeImoPstWriter.Name)
        }, null, "store");
        PstWriterContextResult nameMap = PstPropertyContextWriter.Write(_file,
            _namedProperties.BuildProperties(), 65001, null, null, null,
            Report, "named-properties");
        _nodes.Add(new PstWriterNode(0x61, 0, nameMap.DataBid, nameMap.SubnodeBid));

        byte[] uid = _providerUid.ToByteArray();
        var storeProperties = new List<MapiProperty> {
            Property(MapiKnownProperties.PidTag.ReplVersionHistory, CreateReplVersionHistory()),
            Property(MapiKnownProperties.PidTag.RecordKey, uid),
            Property(MapiKnownProperties.PidTag.DisplayName, _options.DisplayName),
            Property(MapiKnownProperties.PidTag.ValidFolderMask, 0x89),
            Property(MapiKnownProperties.PidTag.PstPassword, 0)
        };
        storeProperties.AddRange(writerProvenance);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.IpmSubTreeEntryId,
            EmailStoreSpecialFolderKind.IpmSubtree);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.IpmInboxEntryId,
            EmailStoreSpecialFolderKind.Inbox);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.IpmOutboxEntryId,
            EmailStoreSpecialFolderKind.Outbox);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.IpmWastebasketEntryId,
            EmailStoreSpecialFolderKind.DeletedItems);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.IpmSentMailEntryId,
            EmailStoreSpecialFolderKind.SentItems);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.ViewsEntryId,
            EmailStoreSpecialFolderKind.PersonalViews);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.CommonViewsEntryId,
            EmailStoreSpecialFolderKind.CommonViews);
        AddSpecialFolderEntryId(storeProperties, MapiKnownProperties.PidTag.FinderEntryId,
            EmailStoreSpecialFolderKind.SearchRoot);
        PstWriterContextResult store = PstPropertyContextWriter.Write(_file,
            storeProperties, 65001, null, null, null, Report, "store");
        _nodes.Add(new PstWriterNode(0x21, 0, store.DataBid, store.SubnodeBid));

        WriteHierarchyMap();
        ulong emptyQueue = _file.WriteDataTree(Array.Empty<byte>());
        _nodes.Add(new PstWriterNode(0x1E1, 0, emptyQueue));
        var searchActivity = new byte[4];
        PstBinary.WriteUInt32(searchActivity, 0, SpamSearchFolderNid);
        _nodes.Add(new PstWriterNode(0x201, 0, _file.WriteDataTree(searchActivity)));
        _nodes.Add(new PstWriterNode(0xEC1, 0, emptyQueue));
    }

    private void WriteFolder(FolderState folder,
        PstWriterItemJournal.PstWriterItemSortedReader items,
        IReadOnlyDictionary<uint, FolderState[]> childrenByParent,
        bool hasInbox) {
        FolderState[] children = childrenByParent.TryGetValue(folder.Nid, out FolderState[]? nested)
            ? nested
            : Array.Empty<FolderState>();
        var folderProperties = new List<MapiProperty> {
            Property(MapiKnownProperties.PidTag.RecordKey, CreateEntryId(folder.Nid)),
            Property(MapiKnownProperties.PidTag.DisplayName, folder.Name),
            Property(MapiKnownProperties.PidTag.ContentCount, folder.NormalItemCount),
            Property(MapiKnownProperties.PidTag.ContentUnreadCount, folder.UnreadItemCount),
            Property(MapiKnownProperties.PidTag.Subfolders, children.Length > 0),
            Property(MapiKnownProperties.PidTag.AssociatedContentCount, folder.AssociatedItemCount)
        };
        if (!string.IsNullOrWhiteSpace(folder.ContainerClass)) {
            folderProperties.Add(Property(MapiKnownProperties.PidTag.ContainerClass, folder.ContainerClass));
        }
        if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox ||
            (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root && !hasInbox)) {
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmAppointmentEntryId,
                EmailStoreSpecialFolderKind.Calendar);
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmContactEntryId,
                EmailStoreSpecialFolderKind.Contacts);
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmJournalEntryId,
                EmailStoreSpecialFolderKind.Journal);
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmNoteEntryId,
                EmailStoreSpecialFolderKind.Notes);
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmTaskEntryId,
                EmailStoreSpecialFolderKind.Tasks);
            AddSpecialFolderEntryId(folderProperties, MapiKnownProperties.PidTag.IpmDraftsEntryId,
                EmailStoreSpecialFolderKind.Drafts);
        }
        PstWriterContextResult pc = PstPropertyContextWriter.Write(_file,
            folderProperties, 65001, null, null, null, Report,
            string.Concat("folder/", FormatId(folder.Nid)));
        _nodes.Add(new PstWriterNode(folder.Nid, folder.ParentNid, pc.DataBid, pc.SubnodeBid));

        if (folder.IsSearchFolder) {
            _nodes.Add(new PstWriterNode((folder.Nid & ~0x1FU) | 0x06, 0, 0));
            PstWriterContextResult criteria = PstPropertyContextWriter.Write(_file,
                new[] { Property(MapiKnownProperties.PidTag.PstSearchCriteriaFlags, 0) },
                65001, null, null, null, Report,
                string.Concat("folder/", FormatId(folder.Nid), "/search-criteria"));
            _nodes.Add(new PstWriterNode((folder.Nid & ~0x1FU) | 0x07, 0,
                criteria.DataBid, criteria.SubnodeBid));
            WriteFolderTable((folder.Nid & ~0x1FU) | 0x10, folder.Nid,
                Array.Empty<PstWriterTableRow>(), SearchContentsColumns, "search-contents");
            return;
        }

        var hierarchyRows = children.Select(child => new PstWriterTableRow(child.Nid,
            CreateHierarchyRowProperties(child, childrenByParent.ContainsKey(child.Nid)))).ToArray();
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0D, folder.Nid,
            hierarchyRows, HierarchyColumns, "hierarchy");
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0E, folder.Nid,
            items.ReadRows(folder.Nid, associated: false),
            ContentsColumns, "contents");
        WriteFolderTable((folder.Nid & ~0x1FU) | 0x0F, folder.Nid,
            items.ReadRows(folder.Nid, associated: true),
            AssociatedColumns, "associated");
    }

    private IReadOnlyList<MapiProperty> CreateHierarchyRowProperties(
        FolderState child, bool hasSubfolders) {
        var properties = new List<MapiProperty> {
            Property(MapiKnownProperties.PidTag.DisplayName, child.Name),
            Property(MapiKnownProperties.PidTag.ContentCount, child.NormalItemCount),
            Property(MapiKnownProperties.PidTag.ContentUnreadCount, child.UnreadItemCount),
            Property(MapiKnownProperties.PidTag.Subfolders, hasSubfolders)
        };
        if (!string.IsNullOrWhiteSpace(child.ContainerClass)) {
            properties.Add(Property(MapiKnownProperties.PidTag.ContainerClass,
                child.ContainerClass));
        }
        if (!child.IsSearchFolder) {
            properties.Add(Property(0x0E30, MapiPropertyType.Binary,
                CreateReplicaId(child.Nid)));
            properties.Add(Property(MapiKnownProperties.PidTag.ReplChangenum,
                checked((long)child.Nid)));
            properties.Add(Property(MapiKnownProperties.PidTag.ReplVersionHistory,
                CreateReplVersionHistory()));
        }
        return properties;
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

    private void WriteReceiveFolderTable(bool hasInbox) {
        uint targetNid = hasInbox
            ? _folders.Values.Single(item =>
                item.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox).Nid
            : RootFolderNid;
        var row = new PstWriterTableRow(1, new[] {
            Property(MapiKnownProperties.PidTag.MessageClass, string.Empty),
            Property(0x6605, MapiPropertyType.Integer32, unchecked((int)targetNid)),
            Property(MapiKnownProperties.PidTag.LtpRowVer, 7)
        });
        PstWriterContextResult table = PstTableContextWriter.Write(_file,
            new[] { row }, 65001, ReceiveFolderColumns, Report, "receive-folders");
        _nodes.Add(new PstWriterNode(0x62B, 0, table.DataBid, table.SubnodeBid));
    }

    private void WriteOutgoingQueueTable() {
        PstWriterContextResult table = PstTableContextWriter.Write(_file,
            Array.Empty<PstWriterTableRow>(), 65001, OutgoingQueueColumns,
            Report, "outgoing-queue");
        _nodes.Add(new PstWriterNode(0x64C, 0, table.DataBid, table.SubnodeBid));
    }

    private void WriteHierarchyMap() {
        HierarchyMapEntry[] entries = _folders.Values
            .Where(item => item.Nid != RootFolderNid && !item.IsSearchFolder)
            .Select(item => new HierarchyMapEntry(CreateReplicaId(item.Nid), item.Nid))
            .OrderBy(item => item.Key, ByteArrayComparer.Instance)
            .ToArray();
        var records = new byte[checked(entries.Length * 20)];
        for (int index = 0; index < entries.Length; index++) {
            int offset = index * 20;
            Buffer.BlockCopy(entries[index].Key, 0, records, offset, 16);
            PstBinary.WriteUInt32(records, offset + 16, entries[index].Nid);
        }
        var heap = new PstWriterHeap(0x9C);
        var rootPointer = new byte[4];
        uint rootPointerHid = heap.Add(rootPointer);
        var header = new byte[8];
        uint headerHid = heap.Add(header);
        PstBinary.WriteUInt32(rootPointer, 0, headerHid);
        PstWriterBth.Complete(heap, header, 16, 4, records);
        ulong dataBid = _file.WriteDataTreeBlocks(heap.Build(rootPointerHid));
        _nodes.Add(new PstWriterNode(0xC01, 0, dataBid));
    }

    private byte[] CreateReplicaId(uint nid) {
        byte[] value = _providerUid.ToByteArray();
        PstBinary.WriteUInt32(value, 0,
            PstBinary.UInt32(value, 0) ^ nid ^ 0xA5A55A5AU);
        return value;
    }

    private byte[] CreateReplVersionHistory() {
        var value = new byte[24];
        PstBinary.WriteUInt32(value, 0, 1);
        Buffer.BlockCopy(_providerUid.ToByteArray(), 0, value, 4, 16);
        PstBinary.WriteUInt32(value, 20, 1);
        return value;
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
        int? flags = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageFlags);
        return !flags.HasValue || (flags.Value & 1) == 0;
    }

    private static MapiProperty Column(MapiPropertyKey key, MapiPropertyType? wireType = null) {
        MapiPropertyType type = wireType ?? key.PreferredType;
        if (key.IsNamed || !key.PropertyId.HasValue || !key.Accepts(type)) {
            throw new ArgumentException("PST table columns require a tagged property key and accepted wire type.",
                nameof(key));
        }
        return new MapiProperty(key.PropertyId.Value, type, null);
    }

    private static MapiProperty Column(ushort propertyId, MapiPropertyType wireType) =>
        new MapiProperty(propertyId, wireType, null);

    private static MapiProperty Property(MapiPropertyKey key, object? value,
        MapiPropertyType? wireType = null) {
        MapiPropertyType type = wireType ?? key.PreferredType;
        if (key.IsNamed || !key.PropertyId.HasValue || !key.Accepts(type)) {
            throw new ArgumentException("PST properties require a tagged property key and accepted wire type.",
                nameof(key));
        }
        if (value != null && !key.ValueType.IsInstanceOfType(value)) {
            throw new ArgumentException(string.Concat("Property ", key.CanonicalName,
                " received an incompatible managed value."), nameof(value));
        }
        return new MapiProperty(key.PropertyId.Value, type, value);
    }

    private static MapiProperty Property(ushort propertyId, MapiPropertyType wireType,
        object? value) => new MapiProperty(propertyId, wireType, value);

    private void AddSpecialFolderEntryId(ICollection<MapiProperty> properties, MapiPropertyKey<byte[]> key,
        EmailStoreSpecialFolderKind kind) {
        FolderState? folder = _folders.Values.FirstOrDefault(item => item.SpecialFolderKind == kind);
        if (folder != null) {
            properties.Add(Property(key, CreateEntryId(folder.Nid)));
        }
    }

    private sealed class HierarchyMapEntry {
        internal HierarchyMapEntry(byte[] key, uint nid) { Key = key; Nid = nid; }
        internal byte[] Key { get; }
        internal uint Nid { get; }
    }

    private sealed class ByteArrayComparer : IComparer<byte[]> {
        internal static ByteArrayComparer Instance { get; } = new ByteArrayComparer();
        public int Compare(byte[]? left, byte[]? right) {
            if (ReferenceEquals(left, right)) return 0;
            if (left == null) return -1;
            if (right == null) return 1;
            int count = Math.Min(left.Length, right.Length);
            for (int index = 0; index < count; index++) {
                int difference = left[index].CompareTo(right[index]);
                if (difference != 0) return difference;
            }
            return left.Length.CompareTo(right.Length);
        }
    }
}
