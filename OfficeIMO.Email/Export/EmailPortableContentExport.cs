namespace OfficeIMO.Email;

/// <summary>Policy for standalone calendar and contact export. Ordinary attachments and mail envelopes are outside this projection.</summary>
public sealed class EmailPortableContentExportOptions {
    /// <summary>Creates a bounded policy. Lossy regeneration requires an explicit Warn or Allow policy.</summary>
    public EmailPortableContentExportOptions(EmailConversionLossPolicy lossPolicy = EmailConversionLossPolicy.Block,
        long maxSourceBytes = 16L * 1024 * 1024, long maxOutputBytes = 16L * 1024 * 1024) {
        if (!Enum.IsDefined(typeof(EmailConversionLossPolicy), lossPolicy)) throw new ArgumentOutOfRangeException(nameof(lossPolicy));
        if (maxSourceBytes <= 0 || maxSourceBytes > 256L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxSourceBytes));
        if (maxOutputBytes <= 0 || maxOutputBytes > 256L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        LossPolicy = lossPolicy; MaxSourceBytes = maxSourceBytes; MaxOutputBytes = maxOutputBytes;
    }
    /// <summary>Known semantic loss policy.</summary>
    public EmailConversionLossPolicy LossPolicy { get; }
    /// <summary>Maximum imported semantic attachment bytes read.</summary>
    public long MaxSourceBytes { get; }
    /// <summary>Maximum generated UTF-8 content-line bytes.</summary>
    public long MaxOutputBytes { get; }
}

/// <summary>A detached portable model, its serialized snapshot, and projection diagnostics.</summary>
public sealed class EmailPortableContentExportResult<TDocument> where TDocument : class {
    private readonly byte[] _bytes;
    internal EmailPortableContentExportResult(TDocument document, byte[] bytes, bool retained,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Document = document; _bytes = bytes; RetainedSemanticSource = retained; Diagnostics = diagnostics;
    }
    /// <summary>Editable standalone model. Editing it does not change the serialized snapshot returned by ToBytes.</summary>
    public TDocument Document { get; }
    /// <summary>Whether unchanged imported content was used, including unmodeled properties.</summary>
    public bool RetainedSemanticSource { get; }
    /// <summary>Decoding and known regeneration losses. This is not a whole-message preservation report.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Returns an independent copy of the UTF-8 serialized export.</summary>
    public byte[] ToBytes() => (byte[])_bytes.Clone();
}

/// <summary>Exports Outlook appointments/tasks and contacts through the canonical iCalendar and vCard codecs.</summary>
public static class EmailPortableContentExport {
    /// <summary>Creates a standalone ICS stream. Unchanged MIME semantic content retains unknown properties and calendar roots.</summary>
    public static EmailPortableContentExportResult<IcsDocument> ToCalendar(EmailDocument document,
        EmailPortableContentExportOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.OutlookItemKind != OutlookItemKind.Appointment && document.OutlookItemKind != OutlookItemKind.Task)
            throw new ArgumentException("Calendar export requires an appointment or task.", nameof(document));
        var policy = options ?? new EmailPortableContentExportOptions();
        var diagnostics = new List<EmailDiagnostic>();
        EmailAttachment? source = IcsCalendarCodec.FindSemanticAttachment(document);
        bool retained = source != null && EmailConversionAnalyzer.HasUnchangedMimeSemanticSource(document);
        cancellationToken.ThrowIfCancellationRequested();
        IcsDocument calendar;
        if (retained) {
            calendar = new IcsDocument(); calendar.Calendars.Clear();
            foreach (var semantic in ReadSemanticSources(document, true, policy, cancellationToken)) {
                EmailAttachmentTextResult text = semantic.Text;
                diagnostics.AddRange(text.Diagnostics);
                foreach (ContentLineComponent root in IcsDocument.Parse(text.Text, ReaderOptions(policy)).Calendars) {
                    if (root.GetFirstProperty("METHOD") == null && semantic.Attachment.ContentTypeParameters.TryGetValue("method", out string? method) && !string.IsNullOrWhiteSpace(method)) {
                        root.AddProperty("METHOD", method!.Trim().ToUpperInvariant());
                        diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_MIME_METHOD_RETAINED",
                            "The effective calendar method was carried from MIME metadata into the standalone calendar.", EmailDiagnosticSeverity.Information, "calendar/METHOD"));
                    }
                    calendar.Calendars.Add(root);
                }
            }
        } else {
            RequireProjection(document, policy, diagnostics);
            calendar = IcsDocument.Load(IcsCalendarCodec.Create(document, policy.MaxOutputBytes, cancellationToken), ReaderOptions(policy, true));
        }
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = calendar.ToBytes(new ContentLineWriterOptions(maxOutputBytes: policy.MaxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailPortableContentExportResult<IcsDocument>(calendar, bytes, retained, diagnostics.AsReadOnly());
    }

    /// <summary>Creates a standalone vCard stream. Distribution lists require their dedicated group model and are rejected.</summary>
    public static EmailPortableContentExportResult<VCardDocument> ToVCard(EmailDocument document,
        EmailPortableContentExportOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.OutlookItemKind != OutlookItemKind.Contact || document.DistributionList != null ||
            document.MessageClass?.StartsWith("IPM.DistList", StringComparison.OrdinalIgnoreCase) == true)
            throw new ArgumentException("Individual vCard export requires a contact.", nameof(document));
        var policy = options ?? new EmailPortableContentExportOptions();
        var diagnostics = new List<EmailDiagnostic>();
        EmailAttachment? source = VCardCodec.FindSemanticAttachment(document);
        bool retained = source != null && EmailConversionAnalyzer.HasUnchangedMimeSemanticSource(document);
        cancellationToken.ThrowIfCancellationRequested();
        VCardDocument cards;
        if (retained) {
            cards = new VCardDocument(); cards.Cards.Clear();
            foreach (var semantic in ReadSemanticSources(document, false, policy, cancellationToken)) {
                EmailAttachmentTextResult text = semantic.Text;
                diagnostics.AddRange(text.Diagnostics);
                foreach (ContentLineComponent root in VCardDocument.Parse(text.Text, ReaderOptions(policy)).Cards) {
                    NormalizeUtf8Charset(root, diagnostics);
                    cards.Cards.Add(root);
                }
            }
        } else {
            RequireProjection(document, policy, diagnostics);
            byte[] content = VCardCodec.CreateAttachment(document, maxOutputBytes: policy.MaxOutputBytes,
                cancellationToken: cancellationToken).Content!;
            cards = VCardDocument.Load(content, ReaderOptions(policy, true));
        }
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = cards.ToBytes(new ContentLineWriterOptions(maxOutputBytes: policy.MaxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailPortableContentExportResult<VCardDocument>(cards, bytes, retained, diagnostics.AsReadOnly());
    }

    private static ContentLineReaderOptions ReaderOptions(EmailPortableContentExportOptions policy, bool generated = false) =>
        new ContentLineReaderOptions(maxInputBytes: generated ? policy.MaxOutputBytes : checked(policy.MaxSourceBytes * 3));

    private static IEnumerable<(EmailAttachmentTextResult Text, EmailAttachment Attachment)> ReadSemanticSources(EmailDocument document, bool calendar,
        EmailPortableContentExportOptions policy, CancellationToken cancellationToken) {
        long bytes = 0; int count = 0;
        foreach (EmailAttachment attachment in document.Attachments) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!attachment.IsProjectedSemanticContent || (calendar
                ? !string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase)
                : !VCardCodec.IsVCardContentType(attachment.ContentType,
                    attachment.ContentTypeParameters.TryGetValue("profile", out string? profile) ? profile : null))) continue;
            if (++count > 1000) throw new EmailLimitExceededException("MaxSemanticParts", count, 1000);
            if (bytes == policy.MaxSourceBytes) throw new EmailLimitExceededException(nameof(policy.MaxSourceBytes), bytes + 1, policy.MaxSourceBytes);
            EmailAttachmentTextResult text = EmailAttachmentTextReader.Read(attachment, policy.MaxSourceBytes - bytes, cancellationToken);
            bytes += text.BytesRead;
            yield return (text, attachment);
        }
    }

    private static void NormalizeUtf8Charset(ContentLineComponent card, List<EmailDiagnostic> diagnostics) {
        foreach (ContentLineProperty property in card.Properties) {
            // Encoded payload spelling remains untouched, so its original charset still applies.
            bool encoded = property.Parameters.Where(parameter => parameter.Name.Equals("ENCODING", StringComparison.OrdinalIgnoreCase))
                .SelectMany(parameter => parameter.Values).Any(value =>
                    value.Equals("QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) || value.Equals("QP", StringComparison.OrdinalIgnoreCase) ||
                    value.Equals("BASE64", StringComparison.OrdinalIgnoreCase) || value.Equals("B", StringComparison.OrdinalIgnoreCase));
            if (encoded) continue;
            bool changed = false;
            foreach (ContentLineParameter parameter in property.Parameters.Where(parameter => parameter.Name.Equals("CHARSET", StringComparison.OrdinalIgnoreCase))) {
                changed |= parameter.Values.Count != 1 || !parameter.Values[0].Equals("utf-8", StringComparison.OrdinalIgnoreCase);
                parameter.Values.Clear(); parameter.Values.Add("utf-8");
            }
            if (changed) diagnostics.Add(new EmailDiagnostic("EMAIL_VCARD_CHARSET_NORMALIZED",
                "The unencoded property charset was updated to match the UTF-8 standalone output.",
                EmailDiagnosticSeverity.Information, "contact/" + property.Name));
        }
    }

    private static void RequireProjection(EmailDocument document, EmailPortableContentExportOptions policy,
        List<EmailDiagnostic> diagnostics) {
        diagnostics.AddRange(EmailConversionAnalyzer.AnalyzePortableContent(document,
            new EmailWriterOptions(policy.LossPolicy, maxOutputBytes: policy.MaxOutputBytes)));
        EmailDiagnostic? stopped = diagnostics.FirstOrDefault(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        if (stopped != null) throw new InvalidOperationException("Portable content export blocked: " + stopped.Code + ".");
    }
}
