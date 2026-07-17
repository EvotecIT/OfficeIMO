using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreReader {
    private static readonly ISet<ushort> SummaryPropertyIds = new HashSet<ushort> {
        0x001A, 0x0037, 0x0E1D, 0x1035,
        0x0039, 0x3007, 0x0E06,
        0x0042, 0x5D02, 0x0065, 0x0064,
        0x0C1A, 0x5D01, 0x0C1F, 0x0C1E,
        0x0017, 0x0026, 0x0E07, 0x0E08, 0x0E17, 0x0E1B,
        0x3FDE, 0x3FFC, 0x3FFD
    };
    private static readonly ISet<ushort> BodyPropertyIds = new HashSet<ushort>(SummaryPropertyIds) {
        0x1000, 0x1009, 0x1013
    };
    private static readonly ISet<ushort> AttachmentMetadataPropertyIds = new HashSet<ushort> {
        0x0E20, 0x3001, 0x3007, 0x3008,
        0x3701, 0x3704, 0x3705, 0x3707, 0x370B, 0x370D, 0x370E,
        0x3712, 0x3713, 0x7FFE,
        0x3FDE, 0x3FFC, 0x3FFD
    };
    private static readonly ISet<ushort> DeferredAttachmentPropertyIds = new HashSet<ushort> { 0x3701 };
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreSessionLifetime _lifetime;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly Dictionary<uint, PstNodeReference> _folderNodes = new Dictionary<uint, PstNodeReference>();
    private readonly List<EmailStoreFolderInfo> _folderInfos = new List<EmailStoreFolderInfo>();
    private CancellationToken _cancellationToken;
    private PstNdbReader? _ndb;
    private PstNamedPropertyMap _namedProperties = PstNamedPropertyMap.Empty;
    private ushort? _headerItemPropertyId;
    private EmailStoreFormat _format;
    private string? _displayName;
    private long _sourceLength;
    private long _totalAttachmentBytes;
    private bool _completeIndexesLoaded;
    private bool _isPasswordProtected;

    internal PstStoreReader(EmailStoreReaderOptions options,
        EmailStoreSessionLifetime? lifetime = null) {
        _options = options;
        _lifetime = lifetime ?? new EmailStoreSessionLifetime();
    }

    internal EmailStoreFormat Format => _format;
    internal string? DisplayName => _displayName;
    internal long SourceLength => _sourceLength;
    internal IReadOnlyList<EmailStoreFolderInfo> Folders => _folderInfos;
    internal IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;
    internal bool IsPasswordProtected => _isPasswordProtected;

    internal EmailStoreStructuralValidationResult ValidateStructure(
        EmailStoreValidationOptions options, CancellationToken cancellationToken) =>
        Ndb.ValidateStructure(options, cancellationToken);

    internal EmailStoreReadResult Read(Stream stream, EmailStoreFormat format,
        CancellationToken cancellationToken) {
        Open(stream, format, loadCompleteIndexes: true, cancellationToken);
        var store = new EmailStore { Format = format, DisplayName = _displayName };
        var folders = new Dictionary<uint, EmailStoreFolder>();
        foreach (EmailStoreFolderInfo info in _folderInfos) {
            uint nid = ParseId(info.Id);
            var folder = new EmailStoreFolder(info.Id, info.ParentId, info.Name,
                info.SpecialFolderKind, info.ClassificationSource,
                info.ContainerClass, info.IsSearchFolder);
            folders.Add(nid, folder);
            store.MutableFolders.Add(folder);
        }

        List<ItemSelection> selections = EnumerateSelections(
                folders.Keys, _options.IncludeAssociatedItems, _options.IncludeOrphanedItems)
            .Take(_options.MaxItemCount == int.MaxValue ? int.MaxValue : _options.MaxItemCount + 1)
            .ToList();
        if (selections.Count > _options.MaxItemCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                selections.Count, _options.MaxItemCount);
        }

        _totalAttachmentBytes = 0;
        foreach (ItemSelection selection in selections) {
            _cancellationToken.ThrowIfCancellationRequested();
            EmailStoreFolder folder = folders[selection.FolderNid];
            EmailStoreItem item = ReadItem(selection.Node, folder.Id, format,
                selection.IsAssociated, selection.IsOrphaned, EmailStoreItemReadOptions.Default,
                selection.Summary);
            if (selection.IsAssociated) folder.MutableAssociatedItems.Add(item);
            else folder.MutableItems.Add(item);
        }
        return new EmailStoreReadResult(store, _diagnostics.ToArray(), stream.Length);
    }

    internal void Open(Stream stream, EmailStoreFormat format, bool loadCompleteIndexes,
        CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        _format = format;
        _sourceLength = stream.Length;
        _displayName = null;
        _totalAttachmentBytes = 0;
        _completeIndexesLoaded = loadCompleteIndexes;
        _diagnostics.Clear();
        _folderNodes.Clear();
        _folderInfos.Clear();
        _namedProperties = PstNamedPropertyMap.Empty;
        _headerItemPropertyId = null;
        _isPasswordProtected = false;

        PstHeader header = PstHeader.Read(stream, format);
        _ndb = new PstNdbReader(stream, header, _options, cancellationToken);
        if (loadCompleteIndexes) Ndb.LoadIndexes();

        if (TryGetNode(0x61, out PstNodeReference? nameToIdNode) && nameToIdNode != null) {
            const string location = "named-properties";
            IReadOnlyList<MapiProperty> mappingProperties = ReadProperties(
                nameToIdNode.DataBid, nameToIdNode.SubnodeBid, location, applyNamedProperties: false);
            _namedProperties = PstNamedPropertyMap.Read(mappingProperties, _diagnostics, location);
            if (_namedProperties.TryGetPropertyId(
                EmailStoreItemContentAvailability.PsetidCommon, 0x8578, out ushort propertyId)) {
                _headerItemPropertyId = propertyId;
            }
        }

        IReadOnlyList<MapiProperty> storeProperties = Array.Empty<MapiProperty>();
        if (TryGetNode(0x21, out PstNodeReference? storeNode) && storeNode != null) {
            storeProperties = ReadProperties(
                storeNode.DataBid, storeNode.SubnodeBid, "store");
            // PidTagPstPassword is a personal-store protection contract. Cached OSTs can reuse
            // the tag for unrelated provider state and are opened through the Outlook profile,
            // not with the legacy PST password checksum.
            if (format == EmailStoreFormat.Pst) {
                PstPassword.Validate(storeProperties, _options);
                _isPasswordProtected = PstPassword.IsProtected(storeProperties);
            }
            _displayName = GetString(storeProperties, 0x3001);
        }

        List<PstNodeReference> folderNodes = EnumerateNodes()
            .Where(node => node.Type == 0x02 || node.Type == 0x03)
            .OrderBy(node => node.Nid)
            .Take(_options.MaxFolderCount + 1)
            .ToList();
        if (folderNodes.Count > _options.MaxFolderCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxFolderCount),
                folderNodes.Count, _options.MaxFolderCount);
        }

        var folderIds = new HashSet<uint>(folderNodes.Select(node => node.Nid));
        var descriptors = new List<PstFolderDescriptor>(folderNodes.Count);
        foreach (PstNodeReference node in folderNodes) {
            _cancellationToken.ThrowIfCancellationRequested();
            string location = string.Concat("folder/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture));
            IReadOnlyList<MapiProperty> properties = ReadProperties(
                node.DataBid, node.SubnodeBid, location);
            string name = GetString(properties, 0x3001) ??
                string.Concat("Folder 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture));
            string? parentId = node.ParentNid != node.Nid && folderIds.Contains(node.ParentNid)
                ? FormatId(node.ParentNid)
                : null;
            int? itemCount = GetInt(properties, 0x3602);
            int? associatedItemCount = GetInt(properties, 0x3617);
            _folderNodes.Add(node.Nid, node);
            descriptors.Add(new PstFolderDescriptor(
                node, name, parentId, itemCount, associatedItemCount,
                GetString(properties, 0x3613), properties));
        }

        PstFolderDescriptor? root = descriptors.FirstOrDefault(item => item.Node.Nid == 0x122);
        PstFolderDescriptor? inbox = descriptors.FirstOrDefault(item =>
            EmailStoreSpecialFolderClassifier.FromDisplayName(item.Name) == EmailStoreSpecialFolderKind.Inbox);
        var specialFolders = new PstSpecialFolderResolver(
            storeProperties, root?.Properties, inbox?.Properties, folderIds);
        foreach (PstFolderDescriptor descriptor in descriptors) {
            EmailStoreSpecialFolderKind specialFolderKind = specialFolders.Resolve(descriptor.Node.Nid);
            _folderInfos.Add(new EmailStoreFolderInfo(
                FormatId(descriptor.Node.Nid), descriptor.ParentId, descriptor.Name,
                descriptor.ItemCount, descriptor.AssociatedItemCount, specialFolderKind,
                specialFolderKind == EmailStoreSpecialFolderKind.Unknown
                    ? EmailStoreFolderClassificationSource.None
                    : EmailStoreFolderClassificationSource.SourceIdentifier,
                descriptor.ContainerClass, descriptor.Node.Type == 0x03));
        }
    }

    internal IEnumerable<EmailStoreItemReference> EnumerateItemReferences(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        IReadOnlyCollection<uint> folderNids = ResolveFolderNids(options);
        int count = 0;
        foreach (ItemSelection selection in EnumerateSelections(folderNids,
            options.IncludeAssociatedItems, options.IncludeOrphanedItems)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++count > options.MaxItems) yield break;
            yield return new EmailStoreItemReference(
                FormatId(selection.Node.Nid), FormatId(selection.FolderNid),
                selection.IsAssociated, selection.IsOrphaned, selection.Summary);
        }
    }

    internal EmailStoreItemSummary ReadReferencedSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        PstNodeReference node = ResolveReferencedNode(reference);
        string location = string.Concat("item/", reference.Id, "/summary");
        IReadOnlyList<MapiProperty> properties = ReadProperties(
            node.DataBid, node.SubnodeBid, location, includedPropertyIds: GetSummaryPropertyIds());
        return CreateSummary(properties, location);
    }

    internal EmailStoreItem ReadReferencedItem(EmailStoreItemReference reference,
        EmailStoreItemReadOptions options, CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        _cancellationToken = cancellationToken;
        PstNodeReference node = ResolveReferencedNode(reference);
        _totalAttachmentBytes = 0;
        return ReadItem(node, reference.FolderId, _format,
            reference.IsAssociated, reference.IsOrphaned, options, reference.Summary);
    }

    private PstNodeReference ResolveReferencedNode(EmailStoreItemReference reference) {
        uint nid = ParseId(reference.Id);
        uint folderNid = ParseId(reference.FolderId);
        if (!_folderNodes.ContainsKey(folderNid) ||
            !TryGetNode(nid, out PstNodeReference? node) ||
            node == null ||
            (node.Type != 0x04 && node.Type != 0x08) ||
            node.ParentNid != folderNid ||
            reference.IsAssociated != (node.Type == 0x08)) {
            throw new KeyNotFoundException("The item reference does not belong to this PST/OST session.");
        }
        return node;
    }

    private IReadOnlyCollection<uint> ResolveFolderNids(EmailStoreEnumerationOptions options) {
        if (options.FolderId == null) return _folderNodes.Keys.ToArray();
        uint requested = ParseId(options.FolderId);
        if (!_folderNodes.ContainsKey(requested)) {
            throw new KeyNotFoundException("The requested folder does not belong to this PST/OST session.");
        }
        var result = new HashSet<uint> { requested };
        if (!options.IncludeDescendants) return result;
        bool added;
        do {
            added = false;
            foreach (PstNodeReference folder in _folderNodes.Values) {
                if (result.Contains(folder.ParentNid) && result.Add(folder.Nid)) added = true;
            }
        } while (added);
        return result;
    }

    private IEnumerable<ItemSelection> EnumerateSelections(IReadOnlyCollection<uint> folderNids,
        bool includeAssociatedItems, bool includeOrphanedItems) {
        var folderSet = new HashSet<uint>(folderNids);
        HashSet<uint>? referenced = includeOrphanedItems ? new HashSet<uint>() : null;
        var foldersWithContents = new HashSet<uint>();
        var foldersWithAssociatedContents = new HashSet<uint>();
        foreach (uint folderNid in folderSet.OrderBy(value => value)) {
            _cancellationToken.ThrowIfCancellationRequested();
            uint nidIndex = folderNid & ~0x1FU;
            uint contentsNid = nidIndex | 0x0EU;
            PstNodeReference folderNode = _folderNodes[folderNid];
            IReadOnlyDictionary<uint, PstSubnodeReference> folderSubnodes =
                Ndb.ReadSubnodes(folderNode.SubnodeBid, _cancellationToken);
            if (TryGetTable(contentsNid, folderSubnodes, out ulong dataBid, out ulong subnodeBid)) {
                foldersWithContents.Add(folderNid);
                foreach (ItemSelection selection in EnumerateTableSelections(
                    dataBid, subnodeBid, folderNid, isAssociated: false)) {
                    referenced?.Add(selection.Node.Nid);
                    yield return selection;
                }
            }

            if (!includeAssociatedItems) continue;
            uint associatedNid = nidIndex | 0x0FU;
            if (TryGetTable(associatedNid, folderSubnodes,
                out ulong associatedDataBid, out ulong associatedSubnodeBid)) {
                foldersWithAssociatedContents.Add(folderNid);
                foreach (ItemSelection selection in EnumerateTableSelections(
                    associatedDataBid, associatedSubnodeBid, folderNid, isAssociated: true)) {
                    referenced?.Add(selection.Node.Nid);
                    yield return selection;
                }
            }
        }

        bool requiresFallbackScan = foldersWithContents.Count < folderSet.Count ||
            (includeAssociatedItems && foldersWithAssociatedContents.Count < folderSet.Count) ||
            includeOrphanedItems;
        if (!requiresFallbackScan) yield break;

        foreach (PstNodeReference node in EnumerateNodes()
            .Where(node => node.Type == 0x04 || (includeAssociatedItems && node.Type == 0x08))) {
            if (!folderSet.Contains(node.ParentNid)) {
                if (!_folderNodes.ContainsKey(node.ParentNid)) AddMissingParentDiagnostic(node);
                continue;
            }

            bool isAssociated = node.Type == 0x08;
            bool hasTable = isAssociated
                ? foldersWithAssociatedContents.Contains(node.ParentNid)
                : foldersWithContents.Contains(node.ParentNid);
            if (!hasTable) {
                yield return new ItemSelection(node, node.ParentNid, isAssociated, isOrphaned: false);
            } else if (includeOrphanedItems && referenced != null && !referenced.Contains(node.Nid)) {
                yield return new ItemSelection(node, node.ParentNid, isAssociated, isOrphaned: true);
            }
        }
    }

    private bool TryGetTable(uint nid, IReadOnlyDictionary<uint, PstSubnodeReference> folderSubnodes,
        out ulong dataBid, out ulong subnodeBid) {
        if (folderSubnodes.TryGetValue(nid, out PstSubnodeReference? subnode)) {
            dataBid = subnode.DataBid;
            subnodeBid = subnode.SubnodeBid;
            return true;
        }
        if (TryGetNode(nid, out PstNodeReference? node) && node != null) {
            dataBid = node.DataBid;
            subnodeBid = node.SubnodeBid;
            return true;
        }
        dataBid = 0;
        subnodeBid = 0;
        return false;
    }

    private IEnumerable<ItemSelection> EnumerateTableSelections(
        ulong dataBid, ulong subnodeBid, uint folderNid, bool isAssociated) {
        string kind = isAssociated ? "associated-contents" : "contents";
        string location = string.Concat("folder/", FormatId(folderNid), "/", kind);
        foreach (IReadOnlyList<MapiProperty> row in EnumerateTableRows(dataBid, subnodeBid, location)) {
            int? rawNid = GetInt(row, 0x67F2);
            if (!rawNid.HasValue) continue;
            uint nid = unchecked((uint)rawNid.Value);
            if (!TryGetNode(nid, out PstNodeReference? itemNode) ||
                itemNode == null ||
                itemNode.Type != (isAssociated ? 0x08 : 0x04)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_CONTENTS_ITEM_MISSING",
                    string.Concat("The ", kind, " table references unavailable item NID 0x",
                        nid.ToString("X", CultureInfo.InvariantCulture), "."),
                    EmailStoreDiagnosticSeverity.Warning,
                    location));
                continue;
            }
            yield return new ItemSelection(
                itemNode, folderNid, isAssociated, isOrphaned: false,
                CreateSummary(row, string.Concat(location, "/item/", FormatId(nid), "/summary")));
        }
    }

    private IEnumerable<IReadOnlyList<MapiProperty>> EnumerateTableRows(
        ulong dataBid, ulong subnodeBid, string location) {
        IEnumerator<IReadOnlyList<MapiProperty>>? rows = null;
        try {
            PstDataTree data = Ndb.OpenDataTree(
                dataBid, _options.MaxDecodedTableBytes, _cancellationToken);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
                Ndb.ReadSubnodes(subnodeBid, _cancellationToken);
            var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
            rows = new PstTableContextReader(
                heap, Ndb.IsUnicode, _options, _cancellationToken,
                message => AddTableCellDiagnostic(message, location)).EnumerateRows().GetEnumerator();
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException) {
            AddTableDiagnostic(exception, location);
        }
        if (rows == null) yield break;

        using (rows) {
            while (true) {
                bool hasRow;
                Exception? failure = null;
                try {
                    hasRow = rows.MoveNext();
                } catch (EmailStoreLimitExceededException) {
                    throw;
                } catch (Exception exception) when (
                    exception is InvalidDataException || exception is NotSupportedException) {
                    hasRow = false;
                    failure = exception;
                }
                if (failure != null) {
                    AddTableDiagnostic(failure, location);
                    yield break;
                }
                if (!hasRow) yield break;
                IReadOnlyList<MapiProperty> row = rows.Current;
                _namedProperties.Apply(row);
                yield return row;
            }
        }
    }

    private void AddTableDiagnostic(Exception exception, string location) {
        _diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_TABLE_CONTEXT",
            exception.Message,
            EmailStoreDiagnosticSeverity.Error,
            location));
    }

    private void AddTableCellDiagnostic(string message, string location) {
        _diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_TABLE_CELL",
            message,
            EmailStoreDiagnosticSeverity.Warning,
            location));
    }

    private void AddMissingParentDiagnostic(PstNodeReference node) {
        _diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_ORPHAN_ITEM",
            string.Concat("Item 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture),
                " references a folder that is not present in the NBT."),
            EmailStoreDiagnosticSeverity.Warning,
            string.Concat("item/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture))));
    }

    private IEnumerable<PstNodeReference> EnumerateNodes() => _completeIndexesLoaded
        ? Ndb.Nodes.Values
        : Ndb.EnumerateNodes(_cancellationToken);

    private bool TryGetNode(uint nid, out PstNodeReference? node) =>
        Ndb.TryGetNode(nid, out node, _cancellationToken);

    private PstNdbReader Ndb => _ndb ?? throw new InvalidOperationException("The NDB has not been initialized.");

    private static string FormatId(uint nid) =>
        string.Concat("pst:", nid.ToString("X8", CultureInfo.InvariantCulture));

    private static uint ParseId(string id) {
        if (id == null || id.Length != 12 || !id.StartsWith("pst:", StringComparison.Ordinal)) {
            throw new ArgumentException("The value is not a valid PST/OST item or folder identifier.", nameof(id));
        }
        uint nid = 0;
        for (int index = 4; index < id.Length; index++) {
            char character = id[index];
            int digit = character >= '0' && character <= '9' ? character - '0'
                : character >= 'A' && character <= 'F' ? character - 'A' + 10
                : character >= 'a' && character <= 'f' ? character - 'a' + 10
                : -1;
            if (digit < 0) {
                throw new ArgumentException(
                    "The value is not a valid PST/OST item or folder identifier.", nameof(id));
            }
            nid = checked((nid << 4) | (uint)digit);
        }
        return nid;
    }

    private static string? GetString(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as string;

    private static int? GetInt(IEnumerable<MapiProperty> properties, ushort id) {
        object? value = properties.FirstOrDefault(property => property.PropertyId == id)?.Value;
        if (value is int number) return number;
        if (value is short shortNumber) return shortNumber;
        return null;
    }

    private static bool? GetBool(IEnumerable<MapiProperty> properties, ushort id) {
        object? value = properties.FirstOrDefault(property => property.PropertyId == id)?.Value;
        if (value is bool boolean) return boolean;
        if (value is int number) return number != 0;
        if (value is short shortNumber) return shortNumber != 0;
        return null;
    }

    private ISet<ushort> GetSummaryPropertyIds() {
        var result = new HashSet<ushort>(SummaryPropertyIds);
        if (_headerItemPropertyId.HasValue) result.Add(_headerItemPropertyId.Value);
        return result;
    }

    private sealed class ItemSelection {
        internal ItemSelection(PstNodeReference node, uint folderNid, bool isAssociated, bool isOrphaned,
            EmailStoreItemSummary? summary = null) {
            Node = node;
            FolderNid = folderNid;
            IsAssociated = isAssociated;
            IsOrphaned = isOrphaned;
            Summary = summary;
        }

        internal PstNodeReference Node { get; }
        internal uint FolderNid { get; }
        internal bool IsAssociated { get; }
        internal bool IsOrphaned { get; }
        internal EmailStoreItemSummary? Summary { get; }
    }

    private sealed class PstFolderDescriptor {
        internal PstFolderDescriptor(PstNodeReference node, string name, string? parentId,
            int? itemCount, int? associatedItemCount, string? containerClass,
            IReadOnlyList<MapiProperty> properties) {
            Node = node;
            Name = name;
            ParentId = parentId;
            ItemCount = itemCount;
            AssociatedItemCount = associatedItemCount;
            ContainerClass = containerClass;
            Properties = properties;
        }

        internal PstNodeReference Node { get; }
        internal string Name { get; }
        internal string? ParentId { get; }
        internal int? ItemCount { get; }
        internal int? AssociatedItemCount { get; }
        internal string? ContainerClass { get; }
        internal IReadOnlyList<MapiProperty> Properties { get; }
    }
}
