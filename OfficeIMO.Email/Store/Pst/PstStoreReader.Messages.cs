using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreReader {
    private EmailStoreItemSummary CreateSummary(IReadOnlyList<MapiProperty> properties, string location) {
        var document = new EmailDocument { Format = EmailFileFormat.Unknown };
        foreach (MapiProperty property in properties) document.MapiProperties.Add(property);
        ProjectItem(document, properties, location);
        int messageStatus = EmailStoreItemContentAvailability.GetMessageStatus(properties) ?? 0;
        return new EmailStoreItemSummary(
            document,
            properties.GetNullableMapiValue(MapiKnownProperties.PidTag.HasAttachments),
            properties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageFlags).HasValue
                ? document.MessageMetadata.IsRead
                : null,
            EmailStoreItemContentAvailability.TryGetHeaderOnly(properties),
            (messageStatus & 0x00001000) != 0,
            (messageStatus & 0x00002000) != 0);
    }

    private EmailStoreItem ReadItem(PstNodeReference node, string folderId, EmailStoreFormat format,
        bool isAssociated, bool isOrphaned, EmailStoreItemReadOptions options,
        EmailStoreItemSummary? summary = null) {
        string id = FormatId(node.Nid);
        string location = string.Concat("item/", id);
        EmailDocument document = ReadItemDocument(
            node.DataBid, node.SubnodeBid, id, folderId, format, location, nestedDepth: 0, options);
        return new EmailStoreItem(id, folderId, document, isAssociated, isOrphaned,
            options.Parts, format, summary);
    }

    private EmailDocument ReadItemDocument(ulong dataBid, ulong subnodeBid, string id, string? folderId,
        EmailStoreFormat format, string location, int nestedDepth, EmailStoreItemReadOptions options,
        PstDecodedObjectBudget? decodedObjectBudget = null) {
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
            Ndb.ReadSubnodes(subnodeBid, _cancellationToken);
        ISet<ushort>? includedPropertyIds = GetIncludedItemPropertyIds(options);
        long maximumDecodedBytes = ResolveMaximumDecodedBytes(options, decodedObjectBudget);
        IReadOnlyList<MapiProperty> properties = ReadProperties(
            dataBid, subnodeBid, location, subnodes, includedPropertyIds: includedPropertyIds,
            maximumDecodedBytes: maximumDecodedBytes);
        decodedObjectBudget?.AddProperties(properties);
        var document = new EmailDocument { Format = EmailFileFormat.Unknown };
        document.Properties["EmailStore:Format"] = format.ToString();
        document.Properties["EmailStore:ItemId"] = id;
        if (folderId != null) document.Properties["EmailStore:FolderId"] = folderId;
        foreach (MapiProperty property in properties) document.MapiProperties.Add(property);

        if (options.Includes(EmailStoreItemReadParts.Recipients)) {
            ReadItemRecipients(document, subnodes, location, decodedObjectBudget);
        }
        if (options.Includes(EmailStoreItemReadParts.AttachmentMetadata)) {
            ReadItemAttachments(document, subnodes, format, location, nestedDepth, options,
                decodedObjectBudget);
        }
        if (options.Parts != EmailStoreItemReadParts.None) {
            ProjectItem(document, properties, location,
                ResolveMaximumDecodedBytes(options, decodedObjectBudget));
            decodedObjectBudget?.AddProjectedBodies(document);
        }
        return document;
    }

    private long ResolveMaximumDecodedBytes(
        EmailStoreItemReadOptions options,
        PstDecodedObjectBudget? decodedObjectBudget) {
        long maximum = options.MaxDecodedPropertyBytes ?? _options.MaxDecodedPropertyBytesPerItem;
        return decodedObjectBudget == null ? maximum : Math.Min(maximum, decodedObjectBudget.RemainingBytes);
    }

    private ISet<ushort>? GetIncludedItemPropertyIds(EmailStoreItemReadOptions options) {
        if (options.Includes(EmailStoreItemReadParts.ExtendedMapiProperties)) return null;
        ISet<ushort> result = options.Includes(EmailStoreItemReadParts.Bodies)
            ? new HashSet<ushort>(BodyPropertyIds)
            : options.Includes(EmailStoreItemReadParts.Metadata)
                ? new HashSet<ushort>(SummaryPropertyIds)
                : new HashSet<ushort>();
        if (_headerItemPropertyId.HasValue) result.Add(_headerItemPropertyId.Value);
        if (_globalObjectIdPropertyId.HasValue) result.Add(_globalObjectIdPropertyId.Value);
        if (_cleanGlobalObjectIdPropertyId.HasValue) result.Add(_cleanGlobalObjectIdPropertyId.Value);
        if (_taskGlobalIdPropertyId.HasValue) result.Add(_taskGlobalIdPropertyId.Value);
        return result;
    }

    private void ReadItemRecipients(EmailDocument document,
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes, string location,
        PstDecodedObjectBudget? decodedObjectBudget) {
        foreach (PstSubnodeReference recipientTable in subnodes.Values
            .Where(item => item.Type == 0x12).OrderBy(item => item.Nid)) {
            string recipientLocation = string.Concat(location, "/recipients/", FormatId(recipientTable.Nid));
            try {
                foreach (EmailRecipient recipient in ReadRecipients(
                    recipientTable,
                    recipientLocation,
                    decodedObjectBudget?.RemainingBytes)) {
                    decodedObjectBudget?.AddProperties(recipient.MapiProperties);
                    document.Recipients.Add(recipient);
                }
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
        string location, int nestedDepth, EmailStoreItemReadOptions readOptions,
        PstDecodedObjectBudget? decodedObjectBudget) {
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
            ISet<ushort>? includedPropertyIds = readOptions.Includes(
                EmailStoreItemReadParts.ExtendedMapiProperties)
                ? null
                : AttachmentMetadataPropertyIds;
            PstHeap heap = CreateHeap(attachmentNode.DataBid, attachmentNode.SubnodeBid,
                attachmentSubnodes, decodedObjectBudget?.RemainingBytes);
            IReadOnlyList<MapiProperty> attachmentProperties = ReadProperties(
                heap, attachmentLocation, sourceHnids, includedPropertyIds,
                DeferredAttachmentPropertyIds,
                maximumDecodedBytes: decodedObjectBudget?.RemainingBytes);
            decodedObjectBudget?.AddProperties(attachmentProperties);
            long declaredLength = Math.Max(0,
                attachmentProperties.GetNullableMapiValue(MapiKnownProperties.PidTag.AttachSize) ?? 0);
            MapiProperty? attachData = attachmentProperties.GetMapiProperty(
                MapiKnownProperties.PidTag.AttachData);
            if (attachData?.PropertyType == MapiPropertyType.Object &&
                attachData.RawData != null && attachData.RawData.Length >= 8) {
                declaredLength = Math.Max(declaredLength,
                    PstBinary.UInt32(attachData.RawData, 4));
            }
            EmailAttachment attachment = PstAttachmentProjection.Create(
                attachmentProperties, declaredLength);

            if (attachment.MapiAttachMethod == 5 &&
                readOptions.Includes(EmailStoreItemReadParts.EmbeddedItems)) {
                ReserveEmbeddedAttachment(declaredLength);
            }

            if (readOptions.Includes(EmailStoreItemReadParts.AttachmentContent) &&
                attachment.MapiAttachMethod != 5 &&
                sourceHnids.TryGetValue(MapiKnownProperties.PidTag.AttachData.PropertyId!.Value,
                    out uint contentHnid)) {
                ReadAttachmentContent(attachment, attachmentProperties, heap, contentHnid,
                    declaredLength, readOptions.PreferStreamingAttachmentContent,
                    decodedObjectBudget);
            }
            if (readOptions.Includes(EmailStoreItemReadParts.EmbeddedItems)) {
                TryReadEmbeddedMessage(attachment, attachmentSubnodes, sourceHnids,
                    format, attachmentLocation, nestedDepth, readOptions,
                    decodedObjectBudget);
            }
            document.Attachments.Add(attachment);
        }
    }

    private void ReadAttachmentContent(EmailAttachment attachment,
        IReadOnlyList<MapiProperty> properties, PstHeap heap, uint contentHnid,
        long declaredLength, bool preferStreaming,
        PstDecodedObjectBudget? decodedObjectBudget) {
        if (declaredLength > _options.MaxAttachmentBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                declaredLength, _options.MaxAttachmentBytes);
        }
        if (declaredLength > 0) CountAttachmentBytes(declaredLength);

        if (_options.RetainAttachmentContent && !preferStreaming) {
            long maximumContentBytes = decodedObjectBudget == null
                ? _options.MaxAttachmentBytes
                : Math.Min(_options.MaxAttachmentBytes, decodedObjectBudget.RemainingBytes);
            byte[] content = heap.ResolveHnid(contentHnid, maximumContentBytes);
            if (content.LongLength > _options.MaxAttachmentBytes) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                    content.LongLength, _options.MaxAttachmentBytes);
            }
            if (content.LongLength > declaredLength) {
                CountAttachmentBytes(content.LongLength - declaredLength);
            }
            attachment.Content = content;
            decodedObjectBudget?.Add(content.LongLength);
            attachment.Length = content.LongLength;
            MapiProperty? contentProperty = properties.GetMapiProperty(MapiKnownProperties.PidTag.AttachData);
            if (contentProperty != null) {
                contentProperty.Value = content;
                contentProperty.RawData = content;
            }
            return;
        }

        attachment.ContentSource = new PstAttachmentContentSource(
            heap, contentHnid, declaredLength > 0 ? (long?)declaredLength : null,
            _options.MaxAttachmentBytes, _attachmentBudget, _lifetime);
    }

    private void CountAttachmentBytes(long length) {
        _attachmentBudget.Add(length);
    }

    private void ResetAttachmentBudget() {
        _attachmentBudget = new PstAttachmentAggregateBudget(_options.MaxTotalAttachmentBytes);
    }

    private void ReserveEmbeddedAttachment(long declaredLength) {
        if (declaredLength > _options.MaxAttachmentBytes) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                declaredLength, _options.MaxAttachmentBytes);
        }
        if (declaredLength > 0) CountAttachmentBytes(declaredLength);
    }

    private void TryReadEmbeddedMessage(EmailAttachment attachment,
        IReadOnlyDictionary<uint, PstSubnodeReference> attachmentSubnodes,
        IReadOnlyDictionary<ushort, uint> sourceHnids, EmailStoreFormat format,
        string location, int nestedDepth, EmailStoreItemReadOptions readOptions,
        PstDecodedObjectBudget? parentDecodedObjectBudget) {
        if (attachment.MapiAttachMethod != 5 ||
            !sourceHnids.TryGetValue(MapiKnownProperties.PidTag.AttachData.GetStandardPropertyId(), out uint embeddedNid) ||
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
        var decodedObjectBudget = parentDecodedObjectBudget ??
            new PstDecodedObjectBudget(_options.MaxAttachmentBytes);
        long bytesBefore = decodedObjectBudget.ConsumedBytes;
        long maximumDecodedBytes = GetEmbeddedMessageMaximumDecodedBytes(
            readOptions.MaxDecodedPropertyBytes ?? _options.MaxDecodedPropertyBytesPerItem,
            decodedObjectBudget,
            _options.MaxAttachmentBytes);
        var embeddedReadOptions = new EmailStoreItemReadOptions(
            readOptions.Parts, maximumDecodedBytes,
            readOptions.PreferStreamingAttachmentContent);
        EmailDocument embeddedDocument = ReadItemDocument(
            embeddedNode.DataBid, embeddedNode.SubnodeBid, embeddedId, folderId: null,
            format, string.Concat(location, "/embedded/", embeddedId), nestedDepth + 1,
            embeddedReadOptions, decodedObjectBudget);
        long observedBytes = decodedObjectBudget.ConsumedBytes - bytesBefore;
        if (observedBytes > attachment.Length) {
            CountAttachmentBytes(observedBytes - attachment.Length);
            attachment.Length = observedBytes;
        }
        attachment.EmbeddedDocument = embeddedDocument;
    }

    internal static long GetEmbeddedMessageMaximumDecodedBytes(
        long requestedMaximum,
        PstDecodedObjectBudget decodedObjectBudget,
        long maximumAttachmentBytes) {
        if (decodedObjectBudget.RemainingBytes <= 0) {
            long actual = decodedObjectBudget.ConsumedBytes == long.MaxValue
                ? long.MaxValue
                : decodedObjectBudget.ConsumedBytes + 1L;
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                actual,
                maximumAttachmentBytes);
        }
        return Math.Min(requestedMaximum, decodedObjectBudget.RemainingBytes);
    }

    private void ProjectItem(EmailDocument document, IReadOnlyList<MapiProperty> properties,
        string location, long? maximumDecodedBytes = null) {
        int? codePage = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageCodepage) ??
            properties.GetNullableMapiValue(MapiKnownProperties.PidTag.InternetCodepage) ??
            properties.GetNullableMapiValue(MapiKnownProperties.PidTag.CodePageId);
        EmailReadResult projection = EmailMapiProjection.Project(document, codePage, location: location,
            options: EmailStoreMessageReader.CreateOptions(
                _options,
                includeAttachmentContent: false,
                maxDecodedPropertyBytes: maximumDecodedBytes),
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

    private IReadOnlyList<EmailRecipient> ReadRecipients(PstSubnodeReference table, string location,
        long? maximumDecodedBytes = null) {
        long effectiveMaximum = maximumDecodedBytes ?? _options.MaxDecodedPropertyBytesPerItem;
        PstDataTree data = Ndb.ReadDataTree(
            table.DataBid, effectiveMaximum, _cancellationToken);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
            Ndb.ReadSubnodes(table.SubnodeBid, _cancellationToken);
        var heap = new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
        IReadOnlyList<IReadOnlyList<MapiProperty>> rows = new PstTableContextReader(
            heap, Ndb.IsUnicode, _options, _cancellationToken,
            message => AddTableCellDiagnostic(message, location),
            effectiveMaximum).ReadRows();
        var recipients = new List<EmailRecipient>(rows.Count);
        foreach (IReadOnlyList<MapiProperty> row in rows) {
            _namedProperties.Apply(row);
            int recipientType = row.GetNullableMapiValue(MapiKnownProperties.PidTag.RecipientType) ?? 0;
            EmailRecipientKind kind = recipientType == 1 ? EmailRecipientKind.To
                : recipientType == 2 ? EmailRecipientKind.Cc
                : recipientType == 3 ? EmailRecipientKind.Bcc
                : EmailRecipientKind.Unknown;
            string? displayName = row.GetMapiValueOrDefault(MapiKnownProperties.PidTag.DisplayName) ??
                row.GetMapiValueOrDefault(MapiKnownProperties.PidTag.RecipientDisplayName);
            string? address = row.GetMapiValueOrDefault(MapiKnownProperties.PidTag.SmtpAddress) ??
                row.GetMapiValueOrDefault(MapiKnownProperties.PidTag.EmailAddress);
            var emailAddress = new EmailAddress(address, displayName) {
                AddressType = row.GetMapiValueOrDefault(MapiKnownProperties.PidTag.AddressType)
            };
            var recipient = new EmailRecipient(kind, emailAddress) {
                MapiRowId = row.GetNullableMapiValue(MapiKnownProperties.PidTag.RowId) ??
                    row.GetNullableMapiValue(MapiKnownProperties.PidTag.LtpRowId),
                MapiObjectType = row.GetNullableMapiValue(MapiKnownProperties.PidTag.ObjectType),
                MapiDisplayType = row.GetNullableMapiValue(MapiKnownProperties.PidTag.DisplayType),
                MapiDisplayTypeEx = row.GetNullableMapiValue(MapiKnownProperties.PidTag.DisplayTypeEx)
            };
            foreach (MapiProperty property in row) recipient.MapiProperties.Add(property);
            recipients.Add(recipient);
        }
        return recipients;
    }

    private IReadOnlyList<MapiProperty> ReadProperties(ulong dataBid, ulong subnodeBid, string location,
        IReadOnlyDictionary<uint, PstSubnodeReference>? knownSubnodes = null,
        bool applyNamedProperties = true, IDictionary<ushort, uint>? sourceHnids = null,
        ISet<ushort>? includedPropertyIds = null,
        long? maximumDecodedBytes = null) {
        try {
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = knownSubnodes ??
                Ndb.ReadSubnodes(subnodeBid, _cancellationToken);
            PstHeap heap = CreateHeap(dataBid, subnodeBid, subnodes, maximumDecodedBytes);
            return ReadProperties(heap, location, sourceHnids, includedPropertyIds,
                deferredPropertyIds: null, applyNamedProperties: applyNamedProperties,
                maximumDecodedBytes: maximumDecodedBytes);
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

    private PstHeap CreateHeap(ulong dataBid, ulong subnodeBid,
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes,
        long? maximumDecodedBytes = null) {
        PstDataTree data = Ndb.OpenDataTree(
            dataBid,
            maximumDecodedBytes ?? _options.MaxDecodedPropertyBytesPerItem,
            _cancellationToken);
        return new PstHeap(data, subnodes, Ndb, _options, _cancellationToken);
    }

    private IReadOnlyList<MapiProperty> ReadProperties(PstHeap heap, string location,
        IDictionary<ushort, uint>? sourceHnids, ISet<ushort>? includedPropertyIds,
        ISet<ushort>? deferredPropertyIds, bool applyNamedProperties = true,
        long? maximumDecodedBytes = null) {
        try {
            IReadOnlyList<MapiProperty> properties =
                new PstPropertyContextReader(heap, _options, _cancellationToken)
                    .ReadProperties(sourceHnids, includedPropertyIds, deferredPropertyIds,
                        maximumDecodedBytes);
            if (applyNamedProperties) _namedProperties.Apply(properties);
            return properties;
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (
            exception is InvalidDataException || exception is NotSupportedException) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_PROPERTY_CONTEXT",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                location));
            return Array.Empty<MapiProperty>();
        }
    }
}
