using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email;

/// <summary>A local HTML message copy and the attachment files written alongside it.</summary>
public sealed class EmailHtmlExportResult {
    internal EmailHtmlExportResult(string path, string? assetsDirectory, EmailAttachmentExtractionResult? extraction,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Path = path; AssetsDirectory = assetsDirectory; Extraction = extraction; Diagnostics = diagnostics;
    }
    /// <summary>Committed HTML path.</summary>
    public string Path { get; }
    /// <summary>Generated resource directory, or null when no resources were exported.</summary>
    public string? AssetsDirectory { get; }
    /// <summary>Attachment provenance and hashes, when resources were exported.</summary>
    public EmailAttachmentExtractionResult? Extraction { get; }
    /// <summary>Body and unresolved image diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}

/// <summary>A self-contained HTML message copy whose embedded resources are inline <c>data:</c> URIs.</summary>
public sealed class EmailHtmlContent {
    internal EmailHtmlContent(string html, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Html = html; Diagnostics = diagnostics;
    }
    /// <summary>Complete HTML document.</summary>
    public string Html { get; }
    /// <summary>Body and unresolved resource diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>Writes <see cref="Html"/> to a caller-owned stream as UTF-8 without a byte order mark.</summary>
    public void WriteTo(Stream destination) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        byte[] bytes = new UTF8Encoding(false).GetBytes(Html);
        destination.Write(bytes, 0, bytes.Length);
    }
}

/// <summary>Creates safe local HTML copies, saving embedded images without downloading remote resources.</summary>
/// <remarks>The caller selects the output path. Copies retain the body, sender, date, and To and Cc recipients; HTML sanitization does not anonymize message content.</remarks>
public static class EmailHtmlExporter {
    /// <summary>Exports a handle-free message with an operation-scoped full read.</summary>
    public static EmailHtmlExportResult Export(EmailMessage message, string path, bool includeAttachments = false,
        bool overwrite = false, CancellationToken cancellationToken = default) {
        if (message == null) throw new ArgumentNullException(nameof(message));
        return message.UseDocument(d => Export(d, path, includeAttachments, overwrite, cancellationToken), cancellationToken);
    }

    /// <summary>Exports sanitized HTML and embedded resources. Earlier resource files survive later failures.</summary>
    /// <remarks>Every export uses a fresh asset directory, allowing atomic replacement of the HTML without changing an earlier copy's resources.</remarks>
    public static EmailHtmlExportResult Export(EmailDocument document, string path, bool includeAttachments = false,
        bool overwrite = false, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (path == null) throw new ArgumentNullException(nameof(path));
        string output = System.IO.Path.GetFullPath(path);
        if (!overwrite && File.Exists(output)) throw new IOException("The HTML destination already exists.");
        cancellationToken.ThrowIfCancellationRequested();
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);
        var diagnostics = projection.Diagnostics.ToList();
        var selected = new HashSet<int>();
        foreach (EmailBodyResource resource in projection.Resources) {
            int index = document.Attachments.IndexOf(resource.SourceAttachment);
            if (index >= 0) selected.Add(index);
        }
        if (includeAttachments) for (int i = 0; i < document.Attachments.Count; i++) selected.Add(i);
        string? assets = null;
        EmailAttachmentExtractionResult? extraction = null;
        var paths = new Dictionary<int, string>();
        if (selected.Count > 0) {
            assets = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(output)!,
                System.IO.Path.GetFileNameWithoutExtension(output) + ".assets-" + Guid.NewGuid().ToString("N"));
            extraction = EmailAttachmentExtractor.Extract(document, assets,
                new EmailAttachmentExtractionOptions(includeHidden: true, selectedAttachmentIndexes: selected), cancellationToken);
            diagnostics.AddRange(extraction.Diagnostics);
            foreach (EmailAttachmentExtractionEntry entry in extraction.Entries) {
                diagnostics.AddRange(entry.Diagnostics);
                if (entry.OutputPath != null) paths.Add(int.Parse(entry.SourcePath, System.Globalization.CultureInfo.InvariantCulture),
                    Uri.EscapeDataString(System.IO.Path.GetFileName(assets)) + "/" +
                    Uri.EscapeDataString(System.IO.Path.GetFileName(entry.OutputPath)));
            }
            if (extraction.Truncated || extraction.Entries.Any(e => e.OutputPath == null))
                throw new InvalidDataException("HTML resources could not all be exported. Earlier resource files remain; inspect the source's content availability and extraction limits.");
        }
        string body = CreateSafeBody(projection, diagnostics, _ => {
            paths.TryGetValue(document.Attachments.IndexOf(_.SourceAttachment), out string? local);
            return local;
        }, cancellationToken);
        var links = new StringBuilder();
        if (includeAttachments) foreach (KeyValuePair<int, string> item in paths.OrderBy(p => p.Key)) {
            if (document.Attachments[item.Key].IsInline) continue;
            links.Append("<li><a href=\"").Append(WebUtility.HtmlEncode(item.Value)).Append("\">")
                .Append(WebUtility.HtmlEncode(document.Attachments[item.Key].FileName ?? "attachment")).Append("</a></li>");
        }
        byte[] bytes = new UTF8Encoding(false).GetBytes(CreatePage(document, body, links.ToString()));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFileCommit.WriteAtomically(output, stream => stream.Write(bytes, 0, bytes.Length), cancellationToken,
            overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        return new EmailHtmlExportResult(output, assets, extraction, diagnostics.AsReadOnly());
    }

    /// <summary>Creates self-contained HTML for a handle-free message with an operation-scoped full read.</summary>
    public static EmailHtmlContent ExportContent(EmailMessage message, CancellationToken cancellationToken = default) {
        if (message == null) throw new ArgumentNullException(nameof(message));
        return message.UseDocument(d => ExportContent(d, cancellationToken), cancellationToken);
    }

    /// <summary>Creates sanitized, self-contained HTML in memory without filesystem access.</summary>
    /// <remarks>Embedded body resources, such as <c>cid:</c> images, become <c>data:</c> URIs read within the projection budgets.
    /// Regular attachments are not included. Use <see cref="EmailHtmlContent.WriteTo(Stream)"/> to stream the result.</remarks>
    public static EmailHtmlContent ExportContent(EmailDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);
        var diagnostics = projection.Diagnostics.ToList();
        var dataUrls = new Dictionary<EmailBodyResource, string>();
        string body = CreateSafeBody(projection, diagnostics, _ => {
            if (!dataUrls.TryGetValue(_, out string? url)) {
                url = "data:" + DataUrlMediaType(_.ContentType) + ";base64," +
                    Convert.ToBase64String(_.ReadAllBytes(cancellationToken));
                dataUrls.Add(_, url);
            }
            return url;
        }, cancellationToken);
        string html = CreatePage(document, body, string.Empty);
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailHtmlContent(html, diagnostics.AsReadOnly());
    }

    private static string CreateSafeBody(EmailBodyProjectionResult projection, List<EmailDiagnostic> diagnostics,
        Func<EmailBodyResource, string?> resourceUrl, CancellationToken cancellationToken) {
        HtmlUrlPolicy resources = HtmlUrlPolicy.CreateOfficeIMOProfile();
        resources.ResolvedUrlTransform = reference => {
            EmailBodyResource? resource = projection.ResolveResource(reference);
            string? url = resource == null ? null : resourceUrl(resource);
            if (url != null) return url;
            if (reference.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return reference;
            diagnostics.Add(new EmailDiagnostic("EMAIL_HTML_RESOURCE_UNRESOLVED",
                "A resource reference has no saved local resource.", EmailDiagnosticSeverity.Warning, reference));
            return null;
        };
        var htmlOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        htmlOptions.ResourceUrlPolicy = resources;
        HtmlConversionDocument html = HtmlConversionDocument.Parse(projection.Html, htmlOptions, cancellationToken);
        return EmailBodyProjection.CreateSafeEmailHtml(html);
    }

    private static string CreatePage(EmailDocument document, string body, string attachmentLinks) {
        string subject = WebUtility.HtmlEncode(document.Subject ?? "Message");
        var header = new StringBuilder();
        AppendHeaderField(header, "From", document.From?.ToString());
        AppendHeaderField(header, "Date", document.Date?.ToString("dddd, d MMMM yyyy HH:mm zzz", System.Globalization.CultureInfo.InvariantCulture));
        AppendHeaderField(header, "To", FormatRecipients(document, EmailRecipientKind.To));
        AppendHeaderField(header, "Cc", FormatRecipients(document, EmailRecipientKind.Cc));
        return "<!doctype html><html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'\"><title>" +
            subject + "</title></head><body><header><h1>" + subject + "</h1><p>" + header + "</p></header><hr>" +
            body + (attachmentLinks.Length > 0 ? "<hr><h2>Attachments</h2><ul>" + attachmentLinks + "</ul>" : string.Empty) + "</body></html>";
    }

    private static void AppendHeaderField(StringBuilder header, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        if (header.Length > 0) header.Append("<br>");
        header.Append(name).Append(": ").Append(WebUtility.HtmlEncode(value));
    }

    private static string FormatRecipients(EmailDocument document, EmailRecipientKind kind) =>
        string.Join(", ", document.Recipients.Where(_ => _.Kind == kind).Select(_ => _.Address?.ToString()).Where(_ => !string.IsNullOrEmpty(_)));

    private static string DataUrlMediaType(string contentType) {
        int parameters = contentType.IndexOf(';');
        string value = (parameters >= 0 ? contentType.Substring(0, parameters) : contentType).Trim().ToLowerInvariant();
        int slash = value.IndexOf('/');
        if (slash <= 0 || slash == value.Length - 1 || value.IndexOf('/', slash + 1) >= 0) return "application/octet-stream";
        foreach (char c in value) {
            if (c != '/' && !(c >= 'a' && c <= 'z') && !(c >= '0' && c <= '9') && "!#$&^_.+-".IndexOf(c) < 0) return "application/octet-stream";
        }
        return value;
    }
}
