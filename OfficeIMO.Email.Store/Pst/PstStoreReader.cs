using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstStoreReader {
    private readonly EmailStoreReaderOptions _options;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private CancellationToken _cancellationToken;
    private PstNdbReader? _ndb;
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

        var store = new EmailStore { Format = format };
        if (_ndb.Nodes.TryGetValue(0x21, out PstNodeReference? storeNode)) {
            IReadOnlyList<MapiProperty> storeProperties = ReadProperties(storeNode.DataBid, storeNode.SubnodeBid,
                "store");
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

        List<PstNodeReference> messageNodes = _ndb.Nodes.Values
            .Where(node => node.Type == 0x04)
            .OrderBy(node => node.Nid)
            .ToList();
        if (messageNodes.Count > _options.MaxMessageCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxMessageCount),
                messageNodes.Count, _options.MaxMessageCount);
        }

        foreach (PstNodeReference node in messageNodes) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!folders.TryGetValue(node.ParentNid, out EmailStoreFolder? folder)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_ORPHAN_MESSAGE",
                    string.Concat("Message 0x", node.Nid.ToString("X", CultureInfo.InvariantCulture),
                        " references a folder that is not present in the NBT."),
                    EmailStoreDiagnosticSeverity.Warning,
                    string.Concat("message/0x", node.Nid.ToString("X", CultureInfo.InvariantCulture))));
                continue;
            }
            EmailStoreMessage message = ReadMessage(node, folder.Id, format);
            folder.MutableMessages.Add(message);
        }

        return new EmailStoreReadResult(store, _diagnostics.ToArray(), stream.Length);
    }

    private EmailStoreMessage ReadMessage(PstNodeReference node, string folderId, EmailStoreFormat format) {
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
        return new EmailStoreMessage(id, folderId, document);
    }

    private IReadOnlyList<EmailRecipient> ReadRecipients(PstSubnodeReference table) {
        PstDataTree data = Ndb.ReadDataTree(table.DataBid, _options.MaxDecodedPropertyBytesPerItem);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = Ndb.ReadSubnodes(table.SubnodeBid);
        var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
        IReadOnlyList<IReadOnlyList<MapiProperty>> rows = new PstTableContextReader(
            heap, Ndb.IsUnicode, _options, _cancellationToken).ReadRows();
        var recipients = new List<EmailRecipient>(rows.Count);
        foreach (IReadOnlyList<MapiProperty> row in rows) {
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
        IReadOnlyDictionary<uint, PstSubnodeReference>? knownSubnodes = null) {
        try {
            PstDataTree data = Ndb.ReadDataTree(dataBid, _options.MaxDecodedPropertyBytesPerItem);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = knownSubnodes ?? Ndb.ReadSubnodes(subnodeBid);
            var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
            return new PstPropertyContextReader(heap, _options, _cancellationToken).ReadProperties();
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
}
