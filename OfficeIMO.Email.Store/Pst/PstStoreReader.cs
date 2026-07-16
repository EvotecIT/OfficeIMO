using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreReader {
    private readonly EmailStoreReaderOptions _options;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private CancellationToken _cancellationToken;
    private PstNdbReader? _ndb;
    private PstNamedPropertyMap _namedProperties = PstNamedPropertyMap.Empty;
    private long _totalAttachmentBytes;

    internal PstStoreReader(EmailStoreReaderOptions options) {
        _options = options;
    }

    internal EmailStoreReadResult Read(Stream stream, EmailStoreFormat format,
        CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        PstHeader header = PstHeader.Read(stream, format);
        _ndb = new PstNdbReader(stream, header, _options, cancellationToken);
        _ndb.LoadIndexes();

        if (_ndb.Nodes.TryGetValue(0x61, out PstNodeReference? nameToIdNode)) {
            const string location = "named-properties";
            IReadOnlyList<MapiProperty> mappingProperties = ReadProperties(
                nameToIdNode.DataBid, nameToIdNode.SubnodeBid, location, applyNamedProperties: false);
            _namedProperties = PstNamedPropertyMap.Read(mappingProperties, _diagnostics, location);
        }

        var store = new EmailStore { Format = format };
        if (_ndb.Nodes.TryGetValue(0x21, out PstNodeReference? storeNode)) {
            IReadOnlyList<MapiProperty> storeProperties = ReadProperties(storeNode.DataBid, storeNode.SubnodeBid,
                "store");
            PstPassword.Validate(storeProperties, _options);
            store.DisplayName = GetString(storeProperties, 0x3001);
        }

        List<PstNodeReference> folderNodes = _ndb.Nodes.Values
            .Where(node => node.Type == 0x02 || node.Type == 0x03)
            .OrderBy(node => node.Nid)
            .ToList();
        if (folderNodes.Count > _options.MaxFolderCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxFolderCount),
                folderNodes.Count, _options.MaxFolderCount);
        }

        var folders = new Dictionary<uint, EmailStoreFolder>();
        foreach (PstNodeReference node in folderNodes) {
            _cancellationToken.ThrowIfCancellationRequested();
            string location = string.Concat("folder/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture));
            IReadOnlyList<MapiProperty> properties = ReadProperties(node.DataBid, node.SubnodeBid, location);
            string name = GetString(properties, 0x3001) ??
                string.Concat("Folder 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture));
            string? parentId = node.ParentNid != node.Nid && folderNodes.Any(candidate => candidate.Nid == node.ParentNid)
                ? FormatId(node.ParentNid)
                : null;
            var folder = new EmailStoreFolder(FormatId(node.Nid), parentId, name);
            folders.Add(node.Nid, folder);
            store.MutableFolders.Add(folder);
        }

        List<ItemSelection> itemSelections = SelectItems(folders);
        if (itemSelections.Count > _options.MaxItemCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                itemSelections.Count, _options.MaxItemCount);
        }

        foreach (ItemSelection selection in itemSelections) {
            _cancellationToken.ThrowIfCancellationRequested();
            EmailStoreFolder folder = folders[selection.FolderNid];
            EmailStoreItem item = ReadItem(
                selection.Node, folder.Id, format, selection.IsAssociated, selection.IsOrphaned);
            if (selection.IsAssociated) folder.MutableAssociatedItems.Add(item);
            else folder.MutableItems.Add(item);
        }

        return new EmailStoreReadResult(store, _diagnostics.ToArray(), stream.Length);
    }

    private List<ItemSelection> SelectItems(IReadOnlyDictionary<uint, EmailStoreFolder> folders) {
        var selected = new List<ItemSelection>();
        var referenced = new HashSet<uint>();
        bool hasContentsTables = false;
        foreach (uint folderNid in folders.Keys.OrderBy(value => value)) {
            uint nidIndex = folderNid & ~0x1FU;
            uint contentsNid = nidIndex | 0x0EU;
            PstNodeReference folderNode = Ndb.Nodes[folderNid];
            IReadOnlyDictionary<uint, PstSubnodeReference> folderSubnodes =
                Ndb.ReadSubnodes(folderNode.SubnodeBid);
            if (TryGetTable(contentsNid, folderSubnodes, out ulong dataBid, out ulong subnodeBid)) {
                hasContentsTables = true;
                AddTableSelections(dataBid, subnodeBid, folderNid,
                    isAssociated: false, selected, referenced);
            }

            if (_options.IncludeAssociatedItems) {
                uint associatedNid = nidIndex | 0x0FU;
                if (TryGetTable(associatedNid, folderSubnodes,
                    out ulong associatedDataBid, out ulong associatedSubnodeBid)) {
                    AddTableSelections(associatedDataBid, associatedSubnodeBid, folderNid,
                        isAssociated: true, selected, referenced);
                }
            }
        }

        IEnumerable<PstNodeReference> ordinaryNodes = Ndb.Nodes.Values
            .Where(node => node.Type == 0x04)
            .OrderBy(node => node.Nid);
        if (!hasContentsTables) {
            foreach (PstNodeReference node in ordinaryNodes) {
                if (folders.ContainsKey(node.ParentNid)) {
                    selected.Add(new ItemSelection(node, node.ParentNid, isAssociated: false, isOrphaned: false));
                    referenced.Add(node.Nid);
                } else {
                    AddMissingParentDiagnostic(node);
                }
            }
            return selected;
        }

        if (_options.IncludeOrphanedItems) {
            foreach (PstNodeReference node in ordinaryNodes.Where(node => !referenced.Contains(node.Nid))) {
                if (folders.ContainsKey(node.ParentNid)) {
                    selected.Add(new ItemSelection(node, node.ParentNid, isAssociated: false, isOrphaned: true));
                } else {
                    AddMissingParentDiagnostic(node);
                }
            }
        }
        return selected;
    }

    private bool TryGetTable(uint nid, IReadOnlyDictionary<uint, PstSubnodeReference> folderSubnodes,
        out ulong dataBid, out ulong subnodeBid) {
        if (folderSubnodes.TryGetValue(nid, out PstSubnodeReference? subnode)) {
            dataBid = subnode.DataBid;
            subnodeBid = subnode.SubnodeBid;
            return true;
        }
        if (Ndb.Nodes.TryGetValue(nid, out PstNodeReference? node)) {
            dataBid = node.DataBid;
            subnodeBid = node.SubnodeBid;
            return true;
        }
        dataBid = 0;
        subnodeBid = 0;
        return false;
    }

    private void AddTableSelections(ulong dataBid, ulong subnodeBid, uint folderNid, bool isAssociated,
        ICollection<ItemSelection> selected, ISet<uint> referenced) {
        string kind = isAssociated ? "associated-contents" : "contents";
        string location = string.Concat("folder/", FormatId(folderNid), "/", kind);
        IReadOnlyList<IReadOnlyList<MapiProperty>>? rows = ReadTableRows(dataBid, subnodeBid, location);
        if (rows == null) return;
        foreach (IReadOnlyList<MapiProperty> row in rows) {
            int? rawNid = GetInt(row, 0x67F2);
            if (!rawNid.HasValue) continue;
            uint nid = unchecked((uint)rawNid.Value);
            if (!Ndb.Nodes.TryGetValue(nid, out PstNodeReference? itemNode) ||
                itemNode.Type != (isAssociated ? 0x08 : 0x04)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_CONTENTS_ITEM_MISSING",
                    string.Concat("The ", kind, " table references unavailable item NID 0x",
                        nid.ToString("X", CultureInfo.InvariantCulture), "."),
                    EmailStoreDiagnosticSeverity.Warning,
                    location));
                continue;
            }
            selected.Add(new ItemSelection(itemNode, folderNid, isAssociated, isOrphaned: false));
            referenced.Add(nid);
        }
    }

    private IReadOnlyList<IReadOnlyList<MapiProperty>>? ReadTableRows(
        ulong dataBid, ulong subnodeBid, string location) {
        try {
            PstDataTree data = Ndb.ReadDataTree(dataBid, _options.MaxDecodedPropertyBytesPerItem);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = Ndb.ReadSubnodes(subnodeBid);
            var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
            IReadOnlyList<IReadOnlyList<MapiProperty>> rows = new PstTableContextReader(
                heap, Ndb.IsUnicode, _options, _cancellationToken).ReadRows();
            foreach (IReadOnlyList<MapiProperty> row in rows) _namedProperties.Apply(row);
            return rows;
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_TABLE_CONTEXT",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                location));
            return null;
        }
    }

    private void AddMissingParentDiagnostic(PstNodeReference node) {
        _diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_ORPHAN_ITEM",
            string.Concat("Item 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture),
                " references a folder that is not present in the NBT."),
            EmailStoreDiagnosticSeverity.Warning,
            string.Concat("item/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture))));
    }

    private PstNdbReader Ndb => _ndb ?? throw new InvalidOperationException("The NDB has not been initialized.");

    private static string FormatId(uint nid) => string.Concat("pst:", nid.ToString("X8", CultureInfo.InvariantCulture));

    private static string? GetString(IEnumerable<MapiProperty> properties, ushort id) =>
        properties.FirstOrDefault(property => property.PropertyId == id)?.Value as string;

    private static int? GetInt(IEnumerable<MapiProperty> properties, ushort id) {
        object? value = properties.FirstOrDefault(property => property.PropertyId == id)?.Value;
        if (value is int number) return number;
        if (value is short shortNumber) return shortNumber;
        return null;
    }

    private sealed class ItemSelection {
        internal ItemSelection(PstNodeReference node, uint folderNid, bool isAssociated, bool isOrphaned) {
            Node = node;
            FolderNid = folderNid;
            IsAssociated = isAssociated;
            IsOrphaned = isOrphaned;
        }

        internal PstNodeReference Node { get; }
        internal uint FolderNid { get; }
        internal bool IsAssociated { get; }
        internal bool IsOrphaned { get; }
    }
}
