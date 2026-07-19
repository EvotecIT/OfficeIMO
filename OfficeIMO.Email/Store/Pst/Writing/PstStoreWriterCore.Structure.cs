using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore {
    private static readonly MapiProperty[] HierarchyColumns = {
        Column(MapiKnownProperties.PidTag.ReplItemid), Column(MapiKnownProperties.PidTag.ReplChangenum),
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
        Column(MapiKnownProperties.PidTag.ReplItemid), Column(MapiKnownProperties.PidTag.ReplChangenum),
        Column(MapiKnownProperties.PidTag.ReplVersionHistory), Column(MapiKnownProperties.PidTag.ReplFlags),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromVersionhistory),
        Column(MapiKnownProperties.PidTag.ReplCopiedfromItemid),
        Column(MapiKnownProperties.PidTag.ItemTemporaryFlags),
        Column(MapiKnownProperties.PidTag.LastModificationTime),
        Column(MapiKnownProperties.PidTag.SecureSubmitFlags)
    };
    private static readonly MapiProperty[] AssociatedColumns = {
        Column(MapiKnownProperties.PidTag.MessageClass), Column(MapiKnownProperties.PidTag.MessageFlags),
        Column(MapiKnownProperties.PidTag.MessageStatus), Column(MapiKnownProperties.PidTag.DisplayName),
        Column(MapiKnownProperties.PidTag.OfflineAddressBookName),
        Column(MapiKnownProperties.PidTag.SendOutlookRecallReport),
        Column(MapiKnownProperties.PidTag.OfflineAddressBookTruncatedProperties),
        Column(MapiKnownProperties.PidTag.ViewDescriptorFlags), Column(MapiKnownProperties.PidTag.ViewDescriptorLinkTo),
        Column(MapiKnownProperties.PidTag.ViewDescriptorViewFolder)
    };
    private static readonly MapiProperty[] RecipientColumns = {
        Column(MapiKnownProperties.PidTag.RecipientType), Column(MapiKnownProperties.PidTag.Responsibility),
        Column(MapiKnownProperties.PidTag.RecordKey), Column(MapiKnownProperties.PidTag.ObjectType),
        Column(MapiKnownProperties.PidTag.EntryId), Column(MapiKnownProperties.PidTag.DisplayName),
        Column(MapiKnownProperties.PidTag.AddressType), Column(MapiKnownProperties.PidTag.EmailAddress),
        Column(MapiKnownProperties.PidTag.SearchKey), Column(MapiKnownProperties.PidTag.DisplayType),
        Column(MapiKnownProperties.PidTag.DisplayNamePrintable, MapiPropertyType.String8),
        Column(MapiKnownProperties.PidTag.SendRichInfo)
    };
    private static readonly MapiProperty[] AttachmentColumns = {
        Column(MapiKnownProperties.PidTag.AttachSize), Column(MapiKnownProperties.PidTag.AttachFilename),
        Column(MapiKnownProperties.PidTag.AttachMethod), Column(MapiKnownProperties.PidTag.RenderingPosition)
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
                Array.Empty<PstWriterTableRow>(), ContentsColumns, "search-contents");
            return;
        }

        var hierarchyRows = children.Select(child => new PstWriterTableRow(child.Nid,
            new[] {
                Property(MapiKnownProperties.PidTag.DisplayName, child.Name),
                Property(MapiKnownProperties.PidTag.ContentCount, child.NormalItemCount),
                Property(MapiKnownProperties.PidTag.ContentUnreadCount, child.UnreadItemCount),
                Property(MapiKnownProperties.PidTag.Subfolders, childrenByParent.ContainsKey(child.Nid)),
                Property(MapiKnownProperties.PidTag.ContainerClass, child.ContainerClass)
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

    private void AddSpecialFolderEntryId(ICollection<MapiProperty> properties, MapiPropertyKey<byte[]> key,
        EmailStoreSpecialFolderKind kind) {
        FolderState? folder = _folders.Values.FirstOrDefault(item => item.SpecialFolderKind == kind);
        if (folder != null) {
            properties.Add(Property(key, CreateEntryId(folder.Nid)));
        }
    }
}
