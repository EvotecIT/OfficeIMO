using OfficeIMO.ContentSafety;

namespace OfficeIMO.Email;

/// <summary>Field-selected share artifact with the optional body projection evidence.</summary>
public sealed class EmailHtmlShareCopyResult {
    internal EmailHtmlShareCopyResult(EmailShareCopyResult copy, EmailIndexTextResult? body, int removedFindings) {
        Copy = copy; BodySourceKind = body?.SourceKind; BodyWasTruncated = body?.Truncated ?? false;
        BodyDiagnosticCodes = Array.AsReadOnly(body?.Diagnostics.Select(value => value.Code).Distinct().ToArray() ?? Array.Empty<string>());
        RemovedConcealedFindingCount = removedFindings;
    }
    /// <summary>Independent plain-text share artifact and field selection evidence.</summary>
    public EmailShareCopyResult Copy { get; }
    /// <summary>Selected source representation, or null when explicit replacement text bypassed projection.</summary>
    public EmailBodySourceKind? BodySourceKind { get; }
    /// <summary>Whether the indexing character bound omitted source text.</summary>
    public bool BodyWasTruncated { get; }
    /// <summary>Projection diagnostic codes without original body text or finding previews.</summary>
    public IReadOnlyList<string> BodyDiagnosticCodes { get; }
    /// <summary>Exact content-safety findings explicitly selected and removed before body projection.</summary>
    public int RemovedConcealedFindingCount { get; }
}

/// <summary>Projects HTML/RTF mail to bounded text before creating a field-selected independent artifact.</summary>
public static class EmailHtmlShareCopy {
    /// <summary>Creates a plain share copy without retaining original HTML/RTF markup or resource references. Text still requires caller review.</summary>
    public static EmailHtmlShareCopyResult Create(EmailDocument source, EmailShareCopyOptions? options = null,
        EmailIndexTextOptions? bodyOptions = null, OfficeContentCleanupSelection? cleanupSelection = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        var effective = (options ?? new EmailShareCopyOptions()).Clone();
        EmailIndexTextResult? body = null;
        int removed = 0;
        var bodySource = source;
        if (cleanupSelection != null) {
            if (effective.ReplacementBodyText != null || string.IsNullOrWhiteSpace(source.Body.Html))
                throw new ArgumentException("Selected HTML cleanup requires an original HTML body and no explicit replacement body.", nameof(cleanupSelection));
            int maximum = bodyOptions?.MaxSourceChars ?? effective.MaxTextChars;
            var inspection = new OfficeContentSafetyOptions { MaxCharacters = maximum,
                MaxInputBytes = Math.Min(int.MaxValue, maximum * 4L), MaxPreviewCharacters = 1 };
            var cleaned = HtmlContentSafety.RemoveSelected(source.Body.Html!, cleanupSelection, inspection);
            removed = cleaned.Changes.Count;
            bodySource = new EmailDocument();
            bodySource.Body.Html = Encoding.UTF8.GetString(cleaned.Output);
        }
        if (effective.ReplacementBodyText == null) {
            body = EmailIndexText.Create(bodySource, bodyOptions ?? new EmailIndexTextOptions {
                MaxSourceChars = effective.MaxTextChars, MaxTextChars = effective.MaxTextChars });
            effective.ReplacementBodyText = body.SelectedText;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailHtmlShareCopyResult(EmailShareCopy.Create(source, effective, cancellationToken), body, removed);
    }
}
