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

/// <summary>Creates safe local HTML copies, saving embedded images without downloading remote resources.</summary>
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
        HtmlUrlPolicy resources = HtmlUrlPolicy.CreateOfficeIMOProfile();
        resources.ResolvedUrlTransform = reference => {
            EmailBodyResource? resource = projection.ResolveResource(reference);
            int index = resource == null ? -1 : document.Attachments.IndexOf(resource.SourceAttachment);
            if (paths.TryGetValue(index, out string? local)) return local;
            if (reference.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return reference;
            diagnostics.Add(new EmailDiagnostic("EMAIL_HTML_RESOURCE_UNRESOLVED",
                "A resource reference has no saved local resource.", EmailDiagnosticSeverity.Warning, reference));
            return null;
        };
        var htmlOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        htmlOptions.ResourceUrlPolicy = resources;
        HtmlConversionDocument html = HtmlConversionDocument.Parse(projection.Html, htmlOptions, cancellationToken);
        string body = EmailBodyProjection.CreateSafeEmailHtml(html);
        var links = new StringBuilder();
        if (includeAttachments) foreach (KeyValuePair<int, string> item in paths.OrderBy(p => p.Key)) {
            if (document.Attachments[item.Key].IsInline) continue;
            links.Append("<li><a href=\"").Append(WebUtility.HtmlEncode(item.Value)).Append("\">")
                .Append(WebUtility.HtmlEncode(document.Attachments[item.Key].FileName ?? "attachment")).Append("</a></li>");
        }
        string content = "<!doctype html><html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'\"><title>" +
            WebUtility.HtmlEncode(document.Subject ?? "Message") + "</title></head><body><header><h1>" +
            WebUtility.HtmlEncode(document.Subject ?? "Message") + "</h1><p>From: " +
            WebUtility.HtmlEncode(document.From?.ToString()) + "<br>To: " +
            WebUtility.HtmlEncode(string.Join(", ", document.Recipients.Where(r => r.Kind == EmailRecipientKind.To).Select(r => r.Address))) +
            "<br>Date: " + WebUtility.HtmlEncode(document.Date?.ToString("u")) + "</p></header><hr>" +
            body + (links.Length > 0 ? "<hr><h2>Attachments</h2><ul>" + links + "</ul>" : string.Empty) + "</body></html>";
        byte[] bytes = new UTF8Encoding(false).GetBytes(content);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFileCommit.WriteAtomically(output, stream => stream.Write(bytes, 0, bytes.Length), cancellationToken,
            overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        return new EmailHtmlExportResult(output, assets, extraction, diagnostics.AsReadOnly());
    }
}
