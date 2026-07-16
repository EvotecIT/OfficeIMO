using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Searches selected semantic fields with explicit scan, decode, character, result, progress, and resume bounds.
    /// </summary>
    public EmailStoreContentSearchReport SearchContent(
        EmailStoreContentQuery query,
        IProgress<EmailStoreContentSearchProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        ThrowIfDisposed();

        var results = new List<EmailStoreContentSearchResult>();
        var diagnostics = new List<EmailStoreDiagnostic>();
        int scanned = 0;
        int skipped = 0;
        bool stoppedAtItemLimit = false;
        bool stoppedAtResultLimit = false;
        long startOffset = query.ResumeFrom?.ItemOffset ?? 0;
        EmailStoreQuery? filter = query.MetadataFilter;
        int enumerationMaximum = checked((int)Math.Min(
            int.MaxValue, startOffset + query.MaxItemsScanned + 1L));
        var enumeration = new EmailStoreEnumerationOptions(
            filter?.FolderId,
            filter?.IncludeDescendants ?? false,
            filter?.IncludeAssociatedItems ?? false,
            filter?.IncludeOrphanedItems ?? false,
            enumerationMaximum);

        using (IEnumerator<EmailStoreItemReference> references =
            EnumerateItems(enumeration, cancellationToken).GetEnumerator()) {
            for (long skippedOffset = 0; skippedOffset < startOffset; skippedOffset++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!references.MoveNext()) {
                    ReportProgress(progress, scanned, results.Count, skipped);
                    return new EmailStoreContentSearchReport(
                        results, diagnostics, scanned, skipped, false, false, null);
                }
            }

            while (references.MoveNext()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (scanned >= query.MaxItemsScanned) {
                    stoppedAtItemLimit = true;
                    break;
                }

                EmailStoreItemReference reference = references.Current;
                scanned++;
                try {
                    EmailStoreItemSummary summary = ReadSummary(reference, cancellationToken);
                    if (filter != null && !Matches(filter, summary)) {
                        ReportProgressIfDue(query, progress, scanned, results.Count, skipped);
                        continue;
                    }

                    IReadOnlyList<SearchFieldText> fields = ReadSearchFields(
                        reference, summary, query, cancellationToken);
                    if (!TryMatch(fields, query, out EmailStoreContentSearchFields matchedFields,
                        out string? snippet)) {
                        ReportProgressIfDue(query, progress, scanned, results.Count, skipped);
                        continue;
                    }

                    results.Add(new EmailStoreContentSearchResult(
                        reference, summary, matchedFields, snippet));
                    if (results.Count >= query.MaxResults) {
                        stoppedAtResultLimit = references.MoveNext();
                        break;
                    }
                } catch (Exception exception) when (
                    query.ContinueOnItemError &&
                    (exception is InvalidDataException ||
                     exception is NotSupportedException ||
                     exception is KeyNotFoundException ||
                     exception is EmailStoreLimitExceededException)) {
                    skipped++;
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_CONTENT_SEARCH_ITEM_SKIPPED",
                        exception.Message,
                        exception is EmailStoreLimitExceededException
                            ? EmailStoreDiagnosticSeverity.Warning
                            : EmailStoreDiagnosticSeverity.Error,
                        string.Concat("item/", reference.Id)));
                }
                ReportProgressIfDue(query, progress, scanned, results.Count, skipped);
            }
        }

        ReportProgress(progress, scanned, results.Count, skipped);
        EmailStoreContentSearchCheckpoint? next = stoppedAtItemLimit || stoppedAtResultLimit
            ? new EmailStoreContentSearchCheckpoint(startOffset + scanned)
            : null;
        return new EmailStoreContentSearchReport(
            results, diagnostics, scanned, skipped,
            stoppedAtItemLimit, stoppedAtResultLimit, next);
    }

    private IReadOnlyList<SearchFieldText> ReadSearchFields(
        EmailStoreItemReference reference,
        EmailStoreItemSummary summary,
        EmailStoreContentQuery query,
        CancellationToken cancellationToken) {
        var fields = new List<SearchFieldText>();
        int remaining = query.MaxSearchableCharactersPerItem;
        if ((query.Fields & EmailStoreContentSearchFields.Subject) != 0) {
            AddField(fields, EmailStoreContentSearchFields.Subject, summary.Subject, ref remaining);
        }
        if ((query.Fields & EmailStoreContentSearchFields.Sender) != 0) {
            AddAddressField(fields, EmailStoreContentSearchFields.Sender, summary.From, ref remaining);
            AddAddressField(fields, EmailStoreContentSearchFields.Sender, summary.Sender, ref remaining);
        }

        EmailStoreContentSearchFields itemFields = query.Fields &
            (EmailStoreContentSearchFields.Recipients |
             EmailStoreContentSearchFields.Bodies |
             EmailStoreContentSearchFields.AttachmentNames);
        if (itemFields == EmailStoreContentSearchFields.None || remaining <= 0) {
            return fields;
        }

        EmailStoreItemReadParts parts = EmailStoreItemReadParts.Metadata;
        if ((itemFields & EmailStoreContentSearchFields.Bodies) != 0) parts |= EmailStoreItemReadParts.Bodies;
        if ((itemFields & EmailStoreContentSearchFields.Recipients) != 0) parts |= EmailStoreItemReadParts.Recipients;
        if ((itemFields & EmailStoreContentSearchFields.AttachmentNames) != 0) {
            parts |= EmailStoreItemReadParts.AttachmentMetadata;
        }
        EmailStoreItem item = ReadItem(reference,
            new EmailStoreItemReadOptions(parts, query.MaxDecodedPropertyBytesPerItem),
            cancellationToken);
        EmailDocument document = item.Document;

        if ((query.Fields & EmailStoreContentSearchFields.Recipients) != 0) {
            foreach (EmailRecipient recipient in document.Recipients) {
                AddAddressField(fields, EmailStoreContentSearchFields.Recipients,
                    recipient.Address, ref remaining);
            }
        }
        if ((query.Fields & EmailStoreContentSearchFields.TextBody) != 0) {
            AddField(fields, EmailStoreContentSearchFields.TextBody, document.Body.Text, ref remaining);
        }
        if ((query.Fields & EmailStoreContentSearchFields.HtmlBody) != 0 && remaining > 0) {
            string text = EmailStoreSearchText.HtmlToText(document.Body.Html, remaining);
            AddNormalizedField(fields, EmailStoreContentSearchFields.HtmlBody, text, ref remaining);
        }
        if ((query.Fields & EmailStoreContentSearchFields.RtfBody) != 0) {
            string text = EmailStoreSearchText.RtfToText(
                document.Body.Rtf, remaining, cancellationToken);
            AddNormalizedField(fields, EmailStoreContentSearchFields.RtfBody, text, ref remaining);
        }
        if ((query.Fields & EmailStoreContentSearchFields.AttachmentNames) != 0) {
            foreach (EmailAttachment attachment in document.Attachments) {
                AddField(fields, EmailStoreContentSearchFields.AttachmentNames,
                    attachment.FileName, ref remaining);
                AddField(fields, EmailStoreContentSearchFields.AttachmentNames,
                    attachment.ContentId, ref remaining);
                AddField(fields, EmailStoreContentSearchFields.AttachmentNames,
                    attachment.ContentLocation, ref remaining);
            }
        }
        return fields;
    }

    private static bool TryMatch(IReadOnlyList<SearchFieldText> fields,
        EmailStoreContentQuery query,
        out EmailStoreContentSearchFields matchedFields,
        out string? snippet) {
        var matchedTerms = new bool[query.Terms.Count];
        matchedFields = EmailStoreContentSearchFields.None;
        string? snippetSource = null;
        foreach (SearchFieldText field in fields) {
            bool fieldMatched = false;
            for (int termIndex = 0; termIndex < query.Terms.Count; termIndex++) {
                if (field.Text.IndexOf(query.Terms[termIndex], StringComparison.OrdinalIgnoreCase) < 0) continue;
                matchedTerms[termIndex] = true;
                fieldMatched = true;
            }
            if (!fieldMatched) continue;
            matchedFields |= field.Field;
            if (snippetSource == null) snippetSource = field.Text;
        }

        bool match = query.MatchMode == EmailStoreContentMatchMode.AnyTerm
            ? matchedTerms.Any(value => value)
            : matchedTerms.All(value => value);
        snippet = match && snippetSource != null
            ? EmailStoreSearchText.CreateSnippet(
                snippetSource, query.Terms, query.SnippetCharacters)
            : null;
        return match;
    }

    private static void AddAddressField(List<SearchFieldText> fields,
        EmailStoreContentSearchFields field, EmailAddress? address, ref int remaining) {
        if (address == null) return;
        AddField(fields, field, address.DisplayName, ref remaining);
        AddField(fields, field, address.Address, ref remaining);
        AddField(fields, field, address.RawValue, ref remaining);
    }

    private static void AddField(List<SearchFieldText> fields,
        EmailStoreContentSearchFields field, string? value, ref int remaining) {
        if (remaining <= 0 || string.IsNullOrEmpty(value)) return;
        string normalized = EmailStoreSearchText.Normalize(value, remaining);
        AddNormalizedField(fields, field, normalized, ref remaining);
    }

    private static void AddNormalizedField(List<SearchFieldText> fields,
        EmailStoreContentSearchFields field, string value, ref int remaining) {
        if (remaining <= 0 || value.Length == 0) return;
        string bounded = value.Length <= remaining ? value : value.Substring(0, remaining);
        fields.Add(new SearchFieldText(field, bounded));
        remaining -= bounded.Length;
    }

    private static void ReportProgressIfDue(EmailStoreContentQuery query,
        IProgress<EmailStoreContentSearchProgress>? progress,
        int scanned, int matches, int skipped) {
        if (scanned % query.ProgressInterval == 0) {
            ReportProgress(progress, scanned, matches, skipped);
        }
    }

    private static void ReportProgress(IProgress<EmailStoreContentSearchProgress>? progress,
        int scanned, int matches, int skipped) =>
        progress?.Report(new EmailStoreContentSearchProgress(scanned, matches, skipped));

    private readonly struct SearchFieldText {
        internal SearchFieldText(EmailStoreContentSearchFields field, string text) {
            Field = field;
            Text = text;
        }
        internal EmailStoreContentSearchFields Field { get; }
        internal string Text { get; }
    }
}
