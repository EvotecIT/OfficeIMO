using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstStoreReader {
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

        List<MessageSelection> messageSelections = SelectMessages(folders);
        if (messageSelections.Count > _options.MaxMessageCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxMessageCount),
                messageSelections.Count, _options.MaxMessageCount);
        }

        foreach (MessageSelection selection in messageSelections) {
            _cancellationToken.ThrowIfCancellationRequested();
            EmailStoreFolder folder = folders[selection.FolderNid];
            EmailStoreMessage message = ReadMessage(
                selection.Node, folder.Id, format, selection.IsAssociated, selection.IsOrphaned);
            if (selection.IsAssociated) folder.MutableAssociatedMessages.Add(message);
            else folder.MutableMessages.Add(message);
        }

        return new EmailStoreReadResult(store, _diagnostics.ToArray(), stream.Length);
    }

    private List<MessageSelection> SelectMessages(IReadOnlyDictionary<uint, EmailStoreFolder> folders) {
        var selected = new List<MessageSelection>();
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

            if (_options.IncludeAssociatedMessages) {
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
                    selected.Add(new MessageSelection(node, node.ParentNid, isAssociated: false, isOrphaned: false));
                    referenced.Add(node.Nid);
                } else {
                    AddMissingParentDiagnostic(node);
                }
            }
            return selected;
        }

        if (_options.IncludeOrphanedMessages) {
            foreach (PstNodeReference node in ordinaryNodes.Where(node => !referenced.Contains(node.Nid))) {
                if (folders.ContainsKey(node.ParentNid)) {
                    selected.Add(new MessageSelection(node, node.ParentNid, isAssociated: false, isOrphaned: true));
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
        ICollection<MessageSelection> selected, ISet<uint> referenced) {
        string kind = isAssociated ? "associated-contents" : "contents";
        string location = string.Concat("folder/", FormatId(folderNid), "/", kind);
        IReadOnlyList<IReadOnlyList<MapiProperty>>? rows = ReadTableRows(dataBid, subnodeBid, location);
        if (rows == null) return;
        foreach (IReadOnlyList<MapiProperty> row in rows) {
            int? rawNid = GetInt(row, 0x67F2);
            if (!rawNid.HasValue) continue;
            uint nid = unchecked((uint)rawNid.Value);
            if (!Ndb.Nodes.TryGetValue(nid, out PstNodeReference? messageNode) ||
                messageNode.Type != (isAssociated ? 0x08 : 0x04)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_CONTENTS_ITEM_MISSING",
                    string.Concat("The ", kind, " table references unavailable message NID 0x",
                        nid.ToString("X", CultureInfo.InvariantCulture), "."),
                    EmailStoreDiagnosticSeverity.Warning,
                    location));
                continue;
            }
            selected.Add(new MessageSelection(messageNode, folderNid, isAssociated, isOrphaned: false));
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
            "EMAIL_STORE_PST_ORPHAN_MESSAGE",
            string.Concat("Message 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture),
                " references a folder that is not present in the NBT."),
            EmailStoreDiagnosticSeverity.Warning,
            string.Concat("message/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture))));
    }

    private EmailStoreMessage ReadMessage(PstNodeReference node, string folderId, EmailStoreFormat format,
        bool isAssociated, bool isOrphaned) {
        string id = FormatId(node.Nid);
        string location = string.Concat("message/", id);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = Ndb.ReadSubnodes(node.SubnodeBid);
        IReadOnlyList<MapiProperty> properties = ReadProperties(node.DataBid, node.SubnodeBid, location, subnodes);
        var document = new EmailDocument { Format = EmailFileFormat.Unknown };
        document.Properties["EmailStore:Format"] = format.ToString();
        document.Properties["EmailStore:ItemId"] = id;
        document.Properties["EmailStore:FolderId"] = folderId;
        foreach (MapiProperty property in properties) document.MapiProperties.Add(property);

        foreach (PstSubnodeReference recipientTable in subnodes.Values
            .Where(item => item.Type == 0x12).OrderBy(item => item.Nid)) {
            string recipientLocation = string.Concat(location, "/recipients/", FormatId(recipientTable.Nid));
            try {
                foreach (EmailRecipient recipient in ReadRecipients(recipientTable)) document.Recipients.Add(recipient);
            } catch (EmailStoreLimitExceededException) {
                throw;
            } catch (InvalidDataException exception) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_RECIPIENT_TABLE",
                    exception.Message,
                    EmailStoreDiagnosticSeverity.Error,
                    recipientLocation));
            }
        }

        int attachmentCount = 0;
        foreach (PstSubnodeReference subnode in subnodes.Values.Where(item => item.Type == 0x05).OrderBy(item => item.Nid)) {
            attachmentCount++;
            if (attachmentCount > _options.MaxAttachmentsPerMessage) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentsPerMessage),
                    attachmentCount, _options.MaxAttachmentsPerMessage);
            }
            string attachmentLocation = string.Concat(location, "/attachment/", FormatId(subnode.Nid));
            IReadOnlyList<MapiProperty> attachmentProperties = ReadProperties(
                subnode.DataBid, subnode.SubnodeBid, attachmentLocation);
            document.Attachments.Add(PstAttachmentProjection.Create(
                attachmentProperties, _options, ref _totalAttachmentBytes));
        }

        int? codePage = GetInt(properties, 0x3FFD) ?? GetInt(properties, 0x3FDE) ?? GetInt(properties, 0x3FFC);
        EmailReadResult projection = EmailMapiProjection.Project(document, codePage, location: location,
            cancellationToken: _cancellationToken);
        foreach (EmailDiagnostic diagnostic in projection.Diagnostics) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity == EmailDiagnosticSeverity.Error
                    ? EmailStoreDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Information
                        ? EmailStoreDiagnosticSeverity.Information
                        : EmailStoreDiagnosticSeverity.Warning,
                diagnostic.Location));
        }
        return new EmailStoreMessage(id, folderId, document, isAssociated, isOrphaned);
    }

    private IReadOnlyList<EmailRecipient> ReadRecipients(PstSubnodeReference table) {
        PstDataTree data = Ndb.ReadDataTree(table.DataBid, _options.MaxDecodedPropertyBytesPerItem);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = Ndb.ReadSubnodes(table.SubnodeBid);
        var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
        IReadOnlyList<IReadOnlyList<MapiProperty>> rows = new PstTableContextReader(
            heap, Ndb.IsUnicode, _options, _cancellationToken).ReadRows();
        var recipients = new List<EmailRecipient>(rows.Count);
        foreach (IReadOnlyList<MapiProperty> row in rows) {
            _namedProperties.Apply(row);
            int recipientType = GetInt(row, 0x0C15) ?? 0;
            EmailRecipientKind kind = recipientType == 1 ? EmailRecipientKind.To
                : recipientType == 2 ? EmailRecipientKind.Cc
                : recipientType == 3 ? EmailRecipientKind.Bcc
                : EmailRecipientKind.Unknown;
            string? displayName = GetString(row, 0x3001) ?? GetString(row, 0x5FF6);
            string? address = GetString(row, 0x39FE) ?? GetString(row, 0x3003);
            var emailAddress = new EmailAddress(address, displayName) {
                AddressType = GetString(row, 0x3002)
            };
            var recipient = new EmailRecipient(kind, emailAddress) {
                MapiRowId = GetInt(row, 0x3000) ?? GetInt(row, 0x67F2),
                MapiObjectType = GetInt(row, 0x0FFE),
                MapiDisplayType = GetInt(row, 0x3900),
                MapiDisplayTypeEx = GetInt(row, 0x3905)
            };
            foreach (MapiProperty property in row) recipient.MapiProperties.Add(property);
            recipients.Add(recipient);
        }
        return recipients;
    }

    private IReadOnlyList<MapiProperty> ReadProperties(ulong dataBid, ulong subnodeBid, string location,
        IReadOnlyDictionary<uint, PstSubnodeReference>? knownSubnodes = null,
        bool applyNamedProperties = true) {
        try {
            PstDataTree data = Ndb.ReadDataTree(dataBid, _options.MaxDecodedPropertyBytesPerItem);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = knownSubnodes ?? Ndb.ReadSubnodes(subnodeBid);
            var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
            IReadOnlyList<MapiProperty> properties =
                new PstPropertyContextReader(heap, _options, _cancellationToken).ReadProperties();
            if (applyNamedProperties) _namedProperties.Apply(properties);
            return properties;
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_PROPERTY_CONTEXT",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                location));
            return Array.Empty<MapiProperty>();
        }
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

    private sealed class MessageSelection {
        internal MessageSelection(PstNodeReference node, uint folderNid, bool isAssociated, bool isOrphaned) {
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
