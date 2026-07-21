using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Rtf;
using System.Net;

namespace OfficeIMO.Html;

/// <summary>Thin email image-export bridge over the HTML renderer.</summary>
public static class EmailImageExportExtensions {
    private static readonly Uri FallbackBaseUri =
        new Uri("https://officeimo.invalid/message/");

    /// <summary>Exports one email body surface or selected rendered page.</summary>
    public static OfficeImageExportResult ExportImage(
        this EmailDocument source,
        OfficeImageExportFormat format,
        EmailImageExportOptions? options = null,
        int pageIndex = 0) {
        EmailRenderPreparation preparation = Prepare(source, options);
        OfficeImageExportResult result = preparation.Document.ExportImage(
            format,
            preparation.Options,
            pageIndex);
        return AttachDiagnostics(
            result,
            preparation.Diagnostics,
            preparation.ResultOptions);
    }

    /// <summary>Asynchronously exports one email surface and resolves inline MIME resources.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(
        this EmailDocument source,
        OfficeImageExportFormat format,
        EmailImageExportOptions? options = null,
        int pageIndex = 0,
        CancellationToken cancellationToken = default) {
        EmailRenderPreparation preparation = Prepare(source, options);
        OfficeImageExportResult result = await preparation.Document.ExportImageAsync(
            format,
            preparation.Options,
            pageIndex,
            cancellationToken).ConfigureAwait(false);
        return AttachDiagnostics(
            result,
            preparation.Diagnostics,
            preparation.ResultOptions);
    }

    /// <summary>Exports every rendered email page.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this EmailDocument source,
        OfficeImageExportFormat format,
        EmailImageExportOptions? options = null) {
        var results = new List<OfficeImageExportResult>();
        source.ExportImages(format, results.Add, options);
        return results.AsReadOnly();
    }

    /// <summary>Streams rendered email pages without retaining earlier payloads.</summary>
    public static void ExportImages(
        this EmailDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        EmailRenderPreparation preparation = Prepare(source, options);
        preparation.Document.ExportImages(
            format,
            result => consumer(AttachDiagnostics(
                result,
                preparation.Diagnostics,
                preparation.ResultOptions)),
            preparation.Options,
            cancellationToken);
    }

    /// <summary>Asynchronously exports every rendered email page and resolves inline MIME resources.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(
        this EmailDocument source,
        OfficeImageExportFormat format,
        EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await source.ExportImagesAsync(
            format,
            (result, _) => {
                results.Add(result);
                return Task.CompletedTask;
            },
            options,
            cancellationToken).ConfigureAwait(false);
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously streams rendered email pages and resolves inline MIME resources.</summary>
    public static async Task ExportImagesAsync(
        this EmailDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer,
        EmailImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        EmailRenderPreparation preparation = Prepare(source, options);
        await preparation.Document.ExportImagesAsync(
            format,
            async (result, token) => await consumer(
                AttachDiagnostics(
                    result,
                    preparation.Diagnostics,
                    preparation.ResultOptions),
                token).ConfigureAwait(false),
            preparation.Options,
            cancellationToken).ConfigureAwait(false);
    }

    private static EmailRenderPreparation Prepare(
        EmailDocument source,
        EmailImageExportOptions? options) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        EmailImageExportOptions effective =
            options?.CloneEmail() ?? new EmailImageExportOptions();
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        string body = CreateBodyHtml(source, effective, diagnostics);
        string html = CreateDocumentHtml(source, body, effective);
        effective.BaseUri ??= ResolveBaseUri(source.Body.HtmlContentLocation);
        EmailImageExportOptions renderOptions = effective.CloneEmail();
        renderOptions.Policy = new OfficeImageExportPolicy();
        ConfigureInlineResources(source, renderOptions);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = effective.BaseUri,
                UseBodyContentsOnly = false
            });
        return new EmailRenderPreparation(
            document,
            renderOptions,
            effective,
            diagnostics.AsReadOnly());
    }

    private static string CreateBodyHtml(
        EmailDocument source,
        EmailImageExportOptions options,
        ICollection<OfficeImageExportDiagnostic> diagnostics) {
        if (options.PreferHtmlBody && !string.IsNullOrWhiteSpace(source.Body.Html)) {
            return ExtractHtmlBody(source.Body.Html!);
        }

        if (!string.IsNullOrWhiteSpace(source.Body.Rtf)) {
            try {
                RtfReadResult rtf = RtfDocument.Read(source.Body.Rtf!);
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    "EMAIL_IMAGE_RTF_BODY_PROJECTED",
                    "The email RTF body was projected through the shared RTF-to-HTML adapter.",
                    "Email body",
                    OfficeImageExportLossKind.Approximation));
                return ExtractHtmlBody(rtf.Document.ToHtml());
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is ArgumentException ||
                exception is NotSupportedException) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    "EMAIL_IMAGE_RTF_BODY_UNREADABLE",
                    "The email RTF body could not be projected; plain text was used when available.",
                    "Email body",
                    OfficeImageExportLossKind.Omission));
            }
        }

        if (!string.IsNullOrEmpty(source.Body.Text)) {
            return "<pre class=\"officeimo-email-plain\">" +
                   WebUtility.HtmlEncode(source.Body.Text) +
                   "</pre>";
        }

        if (!options.PreferHtmlBody && !string.IsNullOrWhiteSpace(source.Body.Html)) {
            return ExtractHtmlBody(source.Body.Html!);
        }

        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            "EMAIL_IMAGE_BODY_MISSING",
            "The email does not contain a renderable HTML, RTF, or plain-text body.",
            "Email body",
            OfficeImageExportLossKind.Omission));
        return "<p class=\"officeimo-email-empty\">This message has no renderable body.</p>";
    }

    private static string CreateDocumentHtml(
        EmailDocument source,
        string body,
        EmailImageExportOptions options) {
        var builder = new StringBuilder();
        builder.Append("<!doctype html><html><head><meta charset=\"utf-8\"><style>")
            .Append("html{background:#f4f6f8}body{margin:0;padding:24px;color:#172033;font-family:Arial,sans-serif}")
            .Append(".officeimo-email{max-width:920px;margin:0 auto;background:#fff;border:1px solid #d8dee8;border-radius:10px;padding:28px}")
            .Append(".officeimo-email-header{border-bottom:1px solid #e5e9f0;margin-bottom:24px;padding-bottom:18px}")
            .Append(".officeimo-email-subject{font-size:24px;font-weight:700;margin:0 0 16px}")
            .Append(".officeimo-email-field{display:grid;grid-template-columns:72px 1fr;gap:8px;margin:5px 0;font-size:13px}")
            .Append(".officeimo-email-label{color:#667085;font-weight:600}")
            .Append(".officeimo-email-body{font-size:16px;line-height:1.45;overflow-wrap:anywhere}")
            .Append(".officeimo-email-plain{white-space:pre-wrap;font:inherit;margin:0}")
            .Append(".officeimo-email-empty{color:#667085;font-style:italic}")
            .Append("img{max-width:100%;height:auto}table{max-width:100%}")
            .Append("</style></head><body><main class=\"officeimo-email\">");
        if (options.IncludeMessageHeaders) AppendHeaders(builder, source);
        builder.Append("<section class=\"officeimo-email-body\">")
            .Append(body)
            .Append("</section></main></body></html>");
        return builder.ToString();
    }

    private static void AppendHeaders(StringBuilder builder, EmailDocument source) {
        builder.Append("<header class=\"officeimo-email-header\"><h1 class=\"officeimo-email-subject\">")
            .Append(WebUtility.HtmlEncode(
                string.IsNullOrWhiteSpace(source.Subject)
                    ? "(no subject)"
                    : source.Subject))
            .Append("</h1>");
        AppendField(builder, "From", source.From?.ToString());
        AppendField(builder, "To", JoinRecipients(source, EmailRecipientKind.To));
        AppendField(builder, "Cc", JoinRecipients(source, EmailRecipientKind.Cc));
        AppendField(
            builder,
            "Date",
            source.Date?.ToString("u", System.Globalization.CultureInfo.InvariantCulture));
        builder.Append("</header>");
    }

    private static void AppendField(
        StringBuilder builder,
        string label,
        string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        builder.Append("<div class=\"officeimo-email-field\"><span class=\"officeimo-email-label\">")
            .Append(WebUtility.HtmlEncode(label))
            .Append("</span><span>")
            .Append(WebUtility.HtmlEncode(value))
            .Append("</span></div>");
    }

    private static string? JoinRecipients(
        EmailDocument source,
        EmailRecipientKind kind) {
        string[] recipients = source.Recipients
            .Where(recipient => recipient.Kind == kind)
            .Select(recipient => recipient.Address.ToString())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToArray();
        return recipients.Length == 0 ? null : string.Join(", ", recipients);
    }

    private static string ExtractHtmlBody(string html) {
        IHtmlDocument document = HtmlConversionDocument.ParseSourceDocumentForAnalysis(html);
        string styles = string.Concat(
            document.Head?.QuerySelectorAll("style")
                .Select(element => element.OuterHtml) ??
            Enumerable.Empty<string>());
        return styles + (document.Body?.InnerHtml ??
                         document.DocumentElement?.InnerHtml ??
                         html);
    }

    private static void ConfigureInlineResources(
        EmailDocument source,
        EmailImageExportOptions options) {
        if (!options.IncludeInlineResources) return;
        HtmlUrlPolicy fallbackResourceUrlPolicy =
            (options.ResourceUrlPolicy ?? options.UrlPolicy).Clone();
        HtmlUrlPolicy resourcePolicy = fallbackResourceUrlPolicy.Clone();
        if (resourcePolicy.RestrictUrlSchemes) {
            resourcePolicy.AllowedUrlSchemes.Add("cid");
            resourcePolicy.AllowedUrlSchemes.Add(
                (options.BaseUri ?? FallbackBaseUri).Scheme);
        }
        resourcePolicy.DisallowFileUrls = false;
        options.ResourceUrlPolicy = resourcePolicy;
        HtmlRenderSynchronousResourceResolver? synchronousFallback =
            options.SynchronousResourceResolver;
        options.SynchronousResourceResolver = (
            HtmlRenderResourceRequest request,
            CancellationToken cancellationToken,
            out HtmlResolvedResource? resource) => {
            cancellationToken.ThrowIfCancellationRequested();
            EmailAttachment? attachment = FindAttachment(
                source,
                request,
                options.BaseUri ?? FallbackBaseUri);
            if (attachment != null) {
                byte[]? bytes = ReadAttachment(
                    attachment,
                    options.MaxResourceBytes,
                    cancellationToken);
                resource = bytes is { Length: > 0 }
                    ? new HtmlResolvedResource(
                        bytes,
                        attachment.ContentType ?? "application/octet-stream")
                    : null;
                return true;
            }
            if (request.Uri.Scheme.Equals(
                    "cid",
                    StringComparison.OrdinalIgnoreCase)) {
                resource = null;
                return true;
            }
            if (synchronousFallback != null &&
                HtmlUrlPolicyEvaluator.IsAllowed(
                    request.Uri.AbsoluteUri,
                    fallbackResourceUrlPolicy) &&
                synchronousFallback(
                    request,
                    cancellationToken,
                    out resource)) {
                return true;
            }
            resource = null;
            return false;
        };
        HtmlRenderResourceResolver? fallback = options.ResourceResolver;
        options.ResourceResolver = async (request, cancellationToken) => {
            EmailAttachment? attachment = FindAttachment(
                source,
                request,
                options.BaseUri ?? FallbackBaseUri);
            if (attachment != null) {
                byte[]? bytes = await ReadAttachmentAsync(
                    attachment,
                    options.MaxResourceBytes,
                    cancellationToken).ConfigureAwait(false);
                if (bytes != null && bytes.Length > 0) {
                    return new HtmlResolvedResource(
                        bytes,
                        attachment.ContentType ?? "application/octet-stream");
                }
            }
            if (fallback == null ||
                !HtmlUrlPolicyEvaluator.IsAllowed(
                    request.Uri.AbsoluteUri,
                    fallbackResourceUrlPolicy)) {
                return null;
            }
            return await fallback(request, cancellationToken)
                .ConfigureAwait(false);
        };
    }

    private static EmailAttachment? FindAttachment(
        EmailDocument source,
        HtmlRenderResourceRequest request,
        Uri baseUri) {
        string reference = request.Source.Trim();
        if (request.Uri.Scheme.Equals("cid", StringComparison.OrdinalIgnoreCase)) {
            string contentId = Uri.UnescapeDataString(
                    request.Uri.OriginalString.Substring("cid:".Length))
                .Trim()
                .Trim('<', '>');
            return source.Attachments.FirstOrDefault(attachment =>
                string.Equals(
                    attachment.ContentId?.Trim().Trim('<', '>'),
                    contentId,
                    StringComparison.OrdinalIgnoreCase));
        }

        foreach (EmailAttachment attachment in source.Attachments) {
            if (!string.IsNullOrWhiteSpace(attachment.ContentLocation)) {
                if (string.Equals(
                    attachment.ContentLocation,
                    reference,
                    StringComparison.OrdinalIgnoreCase)) {
                    return attachment;
                }
                if (Uri.TryCreate(
                        baseUri,
                        attachment.ContentLocation,
                        out Uri? resolved) &&
                    string.Equals(
                        resolved.AbsoluteUri,
                        request.Uri.AbsoluteUri,
                        StringComparison.OrdinalIgnoreCase)) {
                    return attachment;
                }
            }
            if (!string.IsNullOrWhiteSpace(attachment.FileName) &&
                string.Equals(
                    attachment.FileName,
                    reference,
                    StringComparison.OrdinalIgnoreCase)) {
                return attachment;
            }
        }
        return null;
    }

    private static byte[]? ReadAttachment(
        EmailAttachment attachment,
        long maximumBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (attachment.Length > maximumBytes) {
            throw new HtmlRenderResourceByteLimitException(
                attachment.Length);
        }
        if (attachment.Content is { Length: > 0 } retained) {
            if (retained.LongLength > maximumBytes) {
                throw new HtmlRenderResourceByteLimitException(
                    retained.LongLength);
            }
            return (byte[])retained.Clone();
        }
        using Stream stream = attachment.OpenContentStream();
        using var output = new MemoryStream();
        byte[] buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (output.Length + read > maximumBytes) {
                throw new HtmlRenderResourceByteLimitException(
                    checked(output.Length + read));
            }
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static async Task<byte[]?> ReadAttachmentAsync(
        EmailAttachment attachment,
        long maximumBytes,
        CancellationToken cancellationToken) {
        if (attachment.Length > maximumBytes) {
            throw new HtmlRenderResourceByteLimitException(
                attachment.Length);
        }
        if (attachment.Content is { Length: > 0 } retained) {
            if (retained.LongLength > maximumBytes) {
                throw new HtmlRenderResourceByteLimitException(
                    retained.LongLength);
            }
            return (byte[])retained.Clone();
        }
        using Stream stream =
            await attachment.OpenContentStreamAsync(cancellationToken)
                .ConfigureAwait(false);
        using var output = new MemoryStream();
        byte[] buffer = new byte[81920];
        while (true) {
            int read = await stream.ReadAsync(
                buffer,
                0,
                buffer.Length,
                cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (output.Length + read > maximumBytes) {
                throw new HtmlRenderResourceByteLimitException(
                    checked(output.Length + read));
            }
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static Uri ResolveBaseUri(string? contentLocation) {
        return !string.IsNullOrWhiteSpace(contentLocation) &&
               Uri.TryCreate(
                   contentLocation,
                   UriKind.Absolute,
                   out Uri? absolute)
            ? absolute
            : FallbackBaseUri;
    }

    private static OfficeImageExportResult AttachDiagnostics(
        OfficeImageExportResult result,
        IReadOnlyList<OfficeImageExportDiagnostic> diagnostics,
        EmailImageExportOptions options) {
        if (diagnostics.Count == 0) return options.EnsureAccepted(result);
        var combined = new List<OfficeImageExportDiagnostic>(
            diagnostics.Count + result.Diagnostics.Count);
        combined.AddRange(diagnostics);
        combined.AddRange(result.Diagnostics);
        return options.EnsureAccepted(new OfficeImageExportResult(
            result.Format,
            result.Width,
            result.Height,
            result.Bytes,
            result.Name,
            result.Source,
            combined,
            result.SavedPath));
    }

    private sealed class EmailRenderPreparation {
        internal EmailRenderPreparation(
            HtmlConversionDocument document,
            EmailImageExportOptions options,
            EmailImageExportOptions resultOptions,
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            Document = document;
            Options = options;
            ResultOptions = resultOptions;
            Diagnostics = diagnostics;
        }

        internal HtmlConversionDocument Document { get; }
        internal EmailImageExportOptions Options { get; }
        internal EmailImageExportOptions ResultOptions { get; }
        internal IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
    }
}
