namespace OfficeIMO.Email;

/// <summary>
/// Projects an already decoded set of MAPI properties onto the common <see cref="EmailDocument"/> model.
/// Container readers such as PST and OLM readers can use this entry point without duplicating MSG semantics.
/// </summary>
public static class EmailMapiProjection {
    /// <summary>
    /// Applies common message fields, typed Outlook item fields, transport headers, and protection metadata to an
    /// existing document. Existing recipients and attachments are preserved.
    /// </summary>
    /// <param name="document">Document whose <see cref="EmailDocument.MapiProperties"/> have been decoded.</param>
    /// <param name="codePage">Preferred code page for legacy MAPI byte strings and HTML.</param>
    /// <param name="options">Reader safety limits used by derived-body projections.</param>
    /// <param name="location">Logical diagnostic location.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The projected document and any structured diagnostics.</returns>
    public static EmailReadResult Project(
        EmailDocument document,
        int? codePage = null,
        EmailReaderOptions? options = null,
        string? location = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var diagnostics = new List<EmailDiagnostic>();
        var state = new MsgParserState(options ?? EmailReaderOptions.Default, diagnostics, cancellationToken);
        var encoding = MapiStringEncodingContext.FromCodePage(codePage);
        string sourceLocation = string.IsNullOrWhiteSpace(location) ? "mapi" : location!;

        document.OutlookCodePage = codePage ?? document.OutlookCodePage ?? encoding.PrimaryCodePage;
        MsgProjection.Apply(document, state, sourceLocation, encoding);
        MsgProjection.ApplyTransportHeaderRecipients(document, state,
            string.Concat(sourceLocation, "/transport-headers"));
        EmailProtectionProjection.Apply(document, diagnostics,
            string.Concat(sourceLocation, "/protection"));

        return new EmailReadResult(document, diagnostics, 0);
    }
}
