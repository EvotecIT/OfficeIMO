namespace OfficeIMO.Email;

/// <summary>Bounded collection exports that retain source calendar scopes and separate contact cards.</summary>
public static class EmailPortableContentCollection {
    /// <summary>Exports selected calendar items into one ICS stream with separate VCALENDAR roots; TZID and METHOD scopes are not merged.</summary>
    public static EmailPortableContentExportResult<IcsDocument> ToCalendars(IEnumerable<EmailDocument> documents,
        EmailPortableContentExportOptions? options = null, int maxItems = 1000, CancellationToken cancellationToken = default) {
        var combined = new IcsDocument(); combined.Calendars.Clear();
        var policy = options ?? new EmailPortableContentExportOptions();
        var diagnostics = new List<EmailDiagnostic>();
        bool retained = Collect(documents, policy, maxItems, diagnostics, cancellationToken, (document, itemPolicy) => {
            var item = EmailPortableContentExport.ToCalendar(document, itemPolicy, cancellationToken);
            foreach (ContentLineComponent calendar in item.Document.Calendars) combined.Calendars.Add(calendar);
            return (item.ToBytes().LongLength, item.RetainedSemanticSource, item.Diagnostics);
        });
        byte[] bytes = combined.ToBytes(new ContentLineWriterOptions(maxOutputBytes: policy.MaxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailPortableContentExportResult<IcsDocument>(combined, bytes, retained, diagnostics.AsReadOnly());
    }

    /// <summary>Exports selected contacts into one ordered VCF stream without deduplicating or replacing any source.</summary>
    public static EmailPortableContentExportResult<VCardDocument> ToVCards(IEnumerable<EmailDocument> documents,
        EmailPortableContentExportOptions? options = null, int maxItems = 1000, CancellationToken cancellationToken = default) {
        var combined = new VCardDocument(); combined.Cards.Clear();
        var policy = options ?? new EmailPortableContentExportOptions();
        var diagnostics = new List<EmailDiagnostic>();
        bool retained = Collect(documents, policy, maxItems, diagnostics, cancellationToken, (document, itemPolicy) => {
            var item = EmailPortableContentExport.ToVCard(document, itemPolicy, cancellationToken);
            foreach (ContentLineComponent card in item.Document.Cards) combined.Cards.Add(card);
            return (item.ToBytes().LongLength, item.RetainedSemanticSource, item.Diagnostics);
        });
        byte[] bytes = combined.ToBytes(new ContentLineWriterOptions(maxOutputBytes: policy.MaxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailPortableContentExportResult<VCardDocument>(combined, bytes, retained, diagnostics.AsReadOnly());
    }

    private static bool Collect(IEnumerable<EmailDocument> documents, EmailPortableContentExportOptions policy, int maxItems,
        List<EmailDiagnostic> diagnostics, CancellationToken cancellationToken,
        Func<EmailDocument, EmailPortableContentExportOptions, (long Bytes, bool Retained, IReadOnlyList<EmailDiagnostic> Diagnostics)> project) {
        if (documents == null) throw new ArgumentNullException(nameof(documents));
        if (maxItems <= 0 || maxItems > 10000) throw new ArgumentOutOfRangeException(nameof(maxItems));
        int count = 0; long bytes = 0; bool retained = true;
        foreach (EmailDocument document in documents) {
            cancellationToken.ThrowIfCancellationRequested();
            if (count == maxItems) throw new EmailLimitExceededException(nameof(maxItems), count + 1, maxItems);
            if (bytes == policy.MaxOutputBytes) throw new EmailLimitExceededException(nameof(policy.MaxOutputBytes), bytes + 1, policy.MaxOutputBytes);
            var item = project(document, new EmailPortableContentExportOptions(policy.LossPolicy,
                policy.MaxSourceBytes, policy.MaxOutputBytes - bytes));
            bytes += item.Bytes; retained &= item.Retained;
            diagnostics.AddRange(item.Diagnostics.Select(diagnostic => new EmailDiagnostic(diagnostic.Code, diagnostic.Message,
                diagnostic.Severity, "items/" + count.ToString(CultureInfo.InvariantCulture) + "/" + diagnostic.Location, diagnostic.LossKind)));
            count++;
        }
        if (count == 0) throw new ArgumentException("At least one portable item is required.", nameof(documents));
        return retained;
    }
}
