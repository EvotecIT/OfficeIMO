using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreReader {
    private EmailStoreItem ReadItem(PstNodeReference node, string folderId, EmailStoreFormat format,
        bool isAssociated, bool isOrphaned) {
        string id = FormatId(node.Nid);
        string location = string.Concat("item/", id);
        EmailDocument document = ReadItemDocument(
            node.DataBid, node.SubnodeBid, id, folderId, format, location, nestedDepth: 0);
        return new EmailStoreItem(id, folderId, document, isAssociated, isOrphaned);
    }

    private EmailDocument ReadItemDocument(ulong dataBid, ulong subnodeBid, string id, string? folderId,
        EmailStoreFormat format, string location, int nestedDepth) {
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
            Ndb.ReadSubnodes(subnodeBid, _cancellationToken);
        IReadOnlyList<MapiProperty> properties = ReadProperties(dataBid, subnodeBid, location, subnodes);
        var document = new EmailDocument { Format = EmailFileFormat.Unknown };
        document.Properties["EmailStore:Format"] = format.ToString();
        document.Properties["EmailStore:ItemId"] = id;
        if (folderId != null) document.Properties["EmailStore:FolderId"] = folderId;
        foreach (MapiProperty property in properties) document.MapiProperties.Add(property);

        ReadItemRecipients(document, subnodes, location);
        ReadItemAttachments(document, subnodes, format, location, nestedDepth);
        ProjectItem(document, properties, location);
        return document;
    }

    private void ReadItemRecipients(EmailDocument document,
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes, string location) {
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
    }

    private void ReadItemAttachments(EmailDocument document,
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes, EmailStoreFormat format,
        string location, int nestedDepth) {
        int attachmentCount = 0;
        foreach (PstSubnodeReference attachmentNode in subnodes.Values
            .Where(item => item.Type == 0x05).OrderBy(item => item.Nid)) {
            attachmentCount++;
            if (attachmentCount > _options.MaxAttachmentsPerItem) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentsPerItem),
                    attachmentCount, _options.MaxAttachmentsPerItem);
            }

            string attachmentLocation = string.Concat(location, "/attachment/", FormatId(attachmentNode.Nid));
            IReadOnlyDictionary<uint, PstSubnodeReference> attachmentSubnodes =
                Ndb.ReadSubnodes(attachmentNode.SubnodeBid, _cancellationToken);
            var sourceHnids = new Dictionary<ushort, uint>();
            IReadOnlyList<MapiProperty> attachmentProperties = ReadProperties(
                attachmentNode.DataBid, attachmentNode.SubnodeBid, attachmentLocation,
                attachmentSubnodes, sourceHnids: sourceHnids);
            EmailAttachment attachment = PstAttachmentProjection.Create(
                attachmentProperties, _options, ref _totalAttachmentBytes);
            TryReadEmbeddedMessage(attachment, attachmentSubnodes, sourceHnids,
                format, attachmentLocation, nestedDepth);
            document.Attachments.Add(attachment);
        }
    }

    private void TryReadEmbeddedMessage(EmailAttachment attachment,
        IReadOnlyDictionary<uint, PstSubnodeReference> attachmentSubnodes,
        IReadOnlyDictionary<ushort, uint> sourceHnids, EmailStoreFormat format,
        string location, int nestedDepth) {
        if (attachment.MapiAttachMethod != 5 ||
            !sourceHnids.TryGetValue(0x3701, out uint embeddedNid) ||
            (embeddedNid & 0x1F) == 0 ||
            !attachmentSubnodes.TryGetValue(embeddedNid, out PstSubnodeReference? embeddedNode)) return;

        if (nestedDepth >= _options.MaxNestedMessageDepth) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_EMBEDDED_DEPTH_LIMIT",
                "An embedded message was preserved as an attachment but not projected because the configured depth limit was reached.",
                EmailStoreDiagnosticSeverity.Warning,
                location));
            return;
        }

        string embeddedId = FormatId(embeddedNode.Nid);
        attachment.EmbeddedDocument = ReadItemDocument(
            embeddedNode.DataBid, embeddedNode.SubnodeBid, embeddedId, folderId: null,
            format, string.Concat(location, "/embedded/", embeddedId), nestedDepth + 1);
    }

    private void ProjectItem(EmailDocument document, IReadOnlyList<MapiProperty> properties, string location) {
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
    }

    private IReadOnlyList<EmailRecipient> ReadRecipients(PstSubnodeReference table) {
        PstDataTree data = Ndb.ReadDataTree(
            table.DataBid, _options.MaxDecodedPropertyBytesPerItem, _cancellationToken);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
            Ndb.ReadSubnodes(table.SubnodeBid, _cancellationToken);
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
        bool applyNamedProperties = true, IDictionary<ushort, uint>? sourceHnids = null) {
        try {
            PstDataTree data = Ndb.ReadDataTree(
                dataBid, _options.MaxDecodedPropertyBytesPerItem, _cancellationToken);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = knownSubnodes ??
                Ndb.ReadSubnodes(subnodeBid, _cancellationToken);
            var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
            IReadOnlyList<MapiProperty> properties =
                new PstPropertyContextReader(heap, _options, _cancellationToken).ReadProperties(sourceHnids);
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
}
