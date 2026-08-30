using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using System.Net;

namespace OfficeIMO.Email;

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
        EmailBodyProjectionResult bodyProjection = EmailBodyProjection.Create(source,
            new EmailBodyProjectionOptions {
                IncludeResources = effective.IncludeInlineResources,
                SelectionPolicy = effective.PreferHtmlBody
                    ? EmailBodySelectionPolicy.Richest
                    : EmailBodySelectionPolicy.RtfFirst,
                RemoteResourcePolicy = effective.RemoteResourcePolicy,
                MaxResourceBytes = effective.MaxResourceBytes,
                MaxResourceCount = effective.MaxInlineResourceCount,
                MaxTotalResourceBytes = effective.MaxTotalInlineResourceBytes,
                BaseUri = effective.BaseUri
            });
        var diagnostics = bodyProjection.Diagnostics.Select(MapBodyDiagnostic).ToList();
        string body = ExtractHtmlBody(bodyProjection.Html);
        string html = CreateDocumentHtml(source, body, effective);
        effective.BaseUri ??= bodyProjection.Document.BaseUri ?? FallbackBaseUri;
        EmailImageExportOptions renderOptions = effective.CloneEmail();
        renderOptions.Policy = new OfficeImageExportPolicy();
        ConfigureInlineResources(bodyProjection, renderOptions);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = effective.BaseUri,
                UrlPolicy = renderOptions.UrlPolicy.Clone(),
                ResourceUrlPolicy = renderOptions.GetResourceUrlPolicy().Clone(),
                UseBodyContentsOnly = false
            });
        return new EmailRenderPreparation(
            document,
            renderOptions,
            effective,
            diagnostics.AsReadOnly());
    }

    private static string CreateDocumentHtml(
        EmailDocument source,
        string body,
        EmailImageExportOptions options) {
        var builder = new StringBuilder();
        builder.Append("<!doctype html><html><head><meta charset=\"utf-8\"><style>")
            .Append("html{background:#f4f6f8}body{margin:0;padding:24px;color:#172033;font-family:\"")
            .Append(EscapeCssString(options.DefaultFontFamily))
            .Append("\",sans-serif}")
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

    private static string EscapeCssString(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (char character in value) {
            switch (character) {
                case '\\': escaped.Append("\\\\"); break;
                case '"': escaped.Append("\\\""); break;
                case '\n': escaped.Append("\\A "); break;
                case '\r': escaped.Append("\\D "); break;
                case '\f': escaped.Append("\\C "); break;
                case '<': escaped.Append("\\3C "); break;
                case '>': escaped.Append("\\3E "); break;
                default:
                    if (character == '\0' || character < ' ' || character == '\u007F') {
                        escaped.Append('\\')
                            .Append(((int)character).ToString("X", System.Globalization.CultureInfo.InvariantCulture))
                            .Append(' ');
                    } else {
                        escaped.Append(character);
                    }
                    break;
            }
        }
        return escaped.ToString();
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
        EmailBodyProjectionResult projection,
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
            EmailBodyResource? attachment = projection.ResolveResource(request.Source, request.Uri);
            if (attachment != null) {
                byte[] bytes;
                try {
                    bytes = attachment.ReadAllBytes(cancellationToken);
                } catch (EmailLimitExceededException exception) {
                    throw MapResourceLimit(exception);
                }
                resource = bytes.Length > 0
                    ? new HtmlResolvedResource(
                        bytes,
                        attachment.ContentType)
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
            EmailBodyResource? attachment = projection.ResolveResource(request.Source, request.Uri);
            if (attachment != null) {
                byte[] bytes;
                try {
                    bytes = await attachment.ReadAllBytesAsync(cancellationToken).ConfigureAwait(false);
                } catch (EmailLimitExceededException exception) {
                    throw MapResourceLimit(exception);
                }
                if (bytes.Length > 0) {
                    return new HtmlResolvedResource(
                        bytes,
                        attachment.ContentType);
                }
                return null;
            }
            if (request.Uri.Scheme.Equals(
                    "cid",
                    StringComparison.OrdinalIgnoreCase)) {
                return null;
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

    private static Exception MapResourceLimit(EmailLimitExceededException exception) =>
        string.Equals(
            exception.LimitName,
            "EmailBodyProjectionOptions.MaxTotalResourceBytes",
            StringComparison.Ordinal)
            ? new HtmlRenderTotalResourceByteLimitException(exception.ActualValue)
            : new HtmlRenderResourceByteLimitException(exception.ActualValue);

    private static OfficeImageExportDiagnostic MapBodyDiagnostic(EmailDiagnostic diagnostic) {
        string code;
        switch (diagnostic.Code) {
            case "EMAIL_BODY_RTF_PROJECTED":
                code = "EMAIL_IMAGE_RTF_BODY_PROJECTED";
                break;
            case "EMAIL_BODY_RTF_UNREADABLE":
                code = "EMAIL_IMAGE_RTF_BODY_UNREADABLE";
                break;
            case "EMAIL_BODY_MISSING":
                code = "EMAIL_IMAGE_BODY_MISSING";
                break;
            default:
                code = diagnostic.Code;
                break;
        }
        return new OfficeImageExportDiagnostic(
            diagnostic.Severity == EmailDiagnosticSeverity.Error
                ? OfficeImageExportDiagnosticSeverity.Error
                : OfficeImageExportDiagnosticSeverity.Warning,
            code,
            diagnostic.Message,
            "Email body",
            diagnostic.Code == "EMAIL_BODY_RTF_PROJECTED"
                ? OfficeConversionLossKind.Approximation
                : OfficeConversionLossKind.Omission);
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
        return options.EnsureAccepted(result.WithDiagnostics(combined));
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
