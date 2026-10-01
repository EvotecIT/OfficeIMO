using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Email;

/// <summary>Offline contact review and explicitly selected vCard consolidation.</summary>
public static class EmailContactConsolidation {
    /// <summary>Snapshots bounded portable fields and groups address candidates. Display names never create a match.</summary>
    public static EmailContactReviewResult Review(IEnumerable<EmailContactSource> sources,
        EmailContactReviewOptions? options = null, OfflineAddressBookIdentityIndex? directory = null,
        CancellationToken cancellationToken = default) {
        if (sources == null) throw new ArgumentNullException(nameof(sources));
        var policy = options ?? new EmailContactReviewOptions();
        var cards = new List<(string Id, ContentLineComponent Card)>();
        var parents = new List<int>(); var identities = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var ids = new HashSet<string>(StringComparer.Ordinal); var diagnostics = new List<EmailDiagnostic>();
        long bytes = 0;
        foreach (EmailContactSource source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cards.Count == policy.MaxContacts) throw new EmailLimitExceededException(nameof(policy.MaxContacts), cards.Count + 1, policy.MaxContacts);
            if (source == null) throw new ArgumentException("A contact source is null.", nameof(sources));
            if (!ids.Add(source.SourceId)) throw new ArgumentException("Contact source IDs must be unique.", nameof(sources));
            if (bytes == policy.MaxTotalBytes)
                throw new EmailLimitExceededException(nameof(policy.MaxTotalBytes), bytes + 1, policy.MaxTotalBytes);
            var export = EmailPortableContentExport.ToVCard(source.Document,
                new EmailPortableContentExportOptions(EmailConversionLossPolicy.Warn,
                    Math.Min(policy.MaxContactBytes, policy.MaxTotalBytes - bytes), Math.Min(policy.MaxContactBytes, policy.MaxTotalBytes - bytes)), cancellationToken);
            bytes += export.ToBytes().LongLength;
            if (export.Document.Cards.Count != 1) throw new ArgumentException("Each contact source must contain exactly one card.", nameof(sources));
            ContentLineComponent card = export.Document.Cards[0];
            if (card.Properties.Count > policy.MaxPropertiesPerCard)
                throw new EmailLimitExceededException(nameof(policy.MaxPropertiesPerCard), card.Properties.Count, policy.MaxPropertiesPerCard);
            diagnostics.AddRange(export.Diagnostics.Select(diagnostic => new EmailDiagnostic(diagnostic.Code,
                diagnostic.Message, diagnostic.Severity, source.SourceId + "/" + diagnostic.Location, diagnostic.LossKind)));
            int index = cards.Count; cards.Add((source.SourceId, card)); parents.Add(index);
            foreach (string identity in GetIdentities(source, card, directory, diagnostics)) {
                if (identities.TryGetValue(identity, out int previous)) parents[Root(index)] = Root(previous);
                else identities.Add(identity, index);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        var groups = cards.Select((card, index) => (card, root: Root(index))).GroupBy(item => item.root)
            .Select(group => BuildGroup(group.Select(item => item.card).ToArray(), cancellationToken)).ToArray();
        return new EmailContactReviewResult(Array.AsReadOnly(groups), diagnostics.AsReadOnly(), policy.MaxContactBytes, policy.MaxPropertiesPerCard);

        int Root(int index) {
            while (parents[index] != index) { parents[index] = parents[parents[index]]; index = parents[index]; }
            return index;
        }
    }

    /// <summary>
    /// Creates one independent vCard from a selected review group. Every conflict needs a choice index.
    /// A choice of -1 explicitly omits a field; required vCard fields remain enforced by validation.
    /// The operation neither replaces source contacts nor silently selects a conflict winner.
    /// </summary>
    public static VCardDocument Consolidate(EmailContactReviewResult review, int groupIndex,
        IReadOnlyDictionary<string, int>? choices = null, CancellationToken cancellationToken = default) {
        if (review == null) throw new ArgumentNullException(nameof(review));
        if (groupIndex < 0 || groupIndex >= review.Groups.Count) throw new ArgumentOutOfRangeException(nameof(groupIndex));
        EmailContactReviewGroup group = review.Groups[groupIndex];
        var fieldKeys = new HashSet<string>(group.Fields.Select(field => field.Key), StringComparer.Ordinal);
        if (choices != null && choices.Keys.Any(key => !fieldKeys.Contains(key)))
            throw new ArgumentException("A selection key does not belong to the chosen group.", nameof(choices));
        var document = new VCardDocument(); ContentLineComponent card = document.Cards[0]; card.Properties.Clear();
        foreach (EmailContactReviewField field in group.Fields) {
            cancellationToken.ThrowIfCancellationRequested();
            bool selected = choices != null && choices.TryGetValue(field.Key, out _);
            int choice = selected ? choices![field.Key] : 0;
            if (!selected && field.HasConflict) throw new InvalidOperationException("A field conflict needs explicit selection: " + field.Key + ".");
            if (choice < -1 || choice >= field.Choices.Count) throw new ArgumentOutOfRangeException(nameof(choices));
            if (choice >= 0) foreach (ContentLineProperty property in field.Choices[choice].CopyProperties()) {
                if (card.Properties.Count == review.MaxProperties)
                    throw new EmailLimitExceededException("MaxPropertiesPerCard", card.Properties.Count + 1, review.MaxProperties);
                card.Properties.Add(property);
            }
        }
        ContentLineValidationIssue? invalid = document.Validate().FirstOrDefault(issue => issue.Severity == ContentLineValidationSeverity.Error);
        if (invalid != null) throw new InvalidOperationException("Selected contact fields are invalid: " + invalid.Code + ".");
        document.ToBytes(new ContentLineWriterOptions(maxOutputBytes: review.MaxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        return document;
    }

    private static IEnumerable<string> GetIdentities(EmailContactSource source, ContentLineComponent card,
        OfflineAddressBookIdentityIndex? directory, List<EmailDiagnostic> diagnostics) {
        var addresses = new List<EmailAddress>();
        OutlookContact? contact = source.Document.Contact;
        OutlookContactEmailAddress[] slots = contact == null ? Array.Empty<OutlookContactEmailAddress>() : new[] { contact.Email1, contact.Email2, contact.Email3 };
        // Use the same decoder as the Outlook vCard projection, including QP and TEXT escaping.
        foreach (ContentLineProperty property in card.GetProperties("EMAIL")) {
            var decoding = new List<EmailDiagnostic>();
            string value = VCardCodec.ReadTextValue(property, decoding, source.SourceId + "/identity");
            diagnostics.AddRange(decoding);
            if (decoding.Any(diagnostic => diagnostic.Severity != EmailDiagnosticSeverity.Information)) continue;
            var address = new EmailAddress(value);
            OutlookContactEmailAddress? typed = slots.FirstOrDefault(slot => slot.Address == value && !string.IsNullOrWhiteSpace(slot.AddressType));
            address.AddressType = typed?.AddressType;
            addresses.Add(address);
        }
        foreach (EmailAddress address in addresses) {
            if (string.IsNullOrWhiteSpace(address.Address) ||
                string.IsNullOrWhiteSpace(OfflineAddressBookIdentityIndex.NormalizeValue(address.Address!))) continue;
            string? direct = EmailSmtpAddress.Normalize(address);
            if (direct != null) yield return "smtp:" + direct;
            if (directory == null) continue;
            OfflineAddressBookIdentityResolution result;
            try {
                result = directory.Resolve(address.Address!, address.AddressType,
                    new OfflineAddressBookIdentityResolutionOptions(allowAccountNameMatch: false, allowDisplayNameMatch: false));
            } catch (ArgumentException) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_CONTACT_IDENTITY_INVALID",
                    "The retained field has no valid directory identity; its original value remains available for review.",
                    EmailDiagnosticSeverity.Information, source.SourceId + "/identity"));
                continue;
            }
            if (result.Candidate?.IsAuthoritativeAddress == true && !result.Candidate.IsDistributionList && result.IndexIsComplete) {
                yield return "directory:" + result.Candidate.Reference.Id;
                string? smtp = EmailSmtpAddress.Normalize(result.Candidate.ToEmailAddress());
                if (smtp != null) yield return "smtp:" + smtp;
            } else if (result.Status == OfflineAddressBookIdentityResolutionStatus.Ambiguous || !result.IndexIsComplete) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_CONTACT_IDENTITY_UNRESOLVED",
                    "The offline identity is ambiguous or the directory index is incomplete; no directory match was used.",
                    EmailDiagnosticSeverity.Warning, source.SourceId + "/identity"));
            }
        }
    }

    private static EmailContactReviewGroup BuildGroup((string Id, ContentLineComponent Card)[] sources, CancellationToken cancellationToken) {
        var fields = new Dictionary<string, Dictionary<string, (ContentLineProperty[] Properties, List<string> Ids)>>(StringComparer.Ordinal);
        foreach (var source in sources) foreach (var properties in source.Card.Properties.GroupBy(FieldKey)) {
            cancellationToken.ThrowIfCancellationRequested();
            string key = properties.Key;
            if (!fields.TryGetValue(key, out var variants)) fields[key] = variants = new Dictionary<string, (ContentLineProperty[], List<string>)>(StringComparer.Ordinal);
            ContentLineProperty[] bundle = properties.ToArray();
            string signature = string.Concat(bundle.Select(property => Part(property.Value)));
            if (variants.TryGetValue(signature, out var existing)) existing.Ids.Add(source.Id);
            else variants.Add(signature, (bundle, new List<string> { source.Id }));
        }
        var result = fields.Select(field => new EmailContactReviewField(field.Key,
            Array.AsReadOnly(field.Value.Values.Select(value => new EmailContactFieldChoice(value.Properties, value.Ids)).ToArray()))).ToArray();
        return new EmailContactReviewGroup(Array.AsReadOnly(sources.Select(source => source.Id).ToArray()), Array.AsReadOnly(result));
    }

    private static string FieldKey(ContentLineProperty property) {
        return Part(property.Group ?? "") + Part(property.Name.ToUpperInvariant()) + string.Concat(property.Parameters
            .Select(parameter => Part(Part(parameter.Name.ToUpperInvariant()) + string.Concat(parameter.Values.Select(Part))))
            .OrderBy(value => value, StringComparer.Ordinal));
    }
    private static string Part(string value) => value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
}
