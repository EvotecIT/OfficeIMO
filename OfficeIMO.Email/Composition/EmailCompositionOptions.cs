namespace OfficeIMO.Email;

/// <summary>Controls artifact composition without authorizing transport or delivery.</summary>
public sealed class EmailCompositionOptions {
    /// <summary>Additional SMTP addresses to exclude from replies, alongside the composing sender.</summary>
    public IList<string> OwnAddresses { get; } = new List<string>();
    /// <summary>Whether to append a plain-text quote. HTML quotation is supplied by the optional HTML bridge.</summary>
    public bool QuoteOriginal { get; set; } = true;
    /// <summary>Maximum quoted UTF-16 code units; truncation preserves Unicode scalar boundaries.</summary>
    public int MaxQuoteChars { get; set; } = 256 * 1024;
    /// <summary>Maximum threading identifiers retained, including the immediate parent.</summary>
    public int MaxReferences { get; set; } = 100;
    /// <summary>Optional draft timestamp; UTC now is used when omitted.</summary>
    public DateTimeOffset? Date { get; set; }
}

/// <summary>A new independent draft and evidence about recipient or quotation choices.</summary>
public sealed class EmailCompositionResult {
    internal EmailCompositionResult(EmailDocument document, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Document = document; Diagnostics = diagnostics;
    }
    /// <summary>Composed draft. Source headers, signatures, Bcc, MAPI metadata and attachments are not copied.</summary>
    public EmailDocument Document { get; }
    /// <summary>Unresolved recipients, missing plain text, and bounded quotation/threading decisions.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}
