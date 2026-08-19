namespace OfficeIMO.Email;

/// <summary>Body source selected for a safe conversion projection.</summary>
public enum EmailBodySourceKind {
    /// <summary>No renderable body was present.</summary>
    None = 0,
    /// <summary>The MIME/Outlook HTML body was selected.</summary>
    Html = 1,
    /// <summary>The RTF body was projected through OfficeIMO.Html.Rtf.</summary>
    Rtf = 2,
    /// <summary>The plain-text body was HTML-encoded.</summary>
    PlainText = 3
}

/// <summary>Deterministic body preference shared by email consumers.</summary>
public enum EmailBodySelectionPolicy {
    /// <summary>Select HTML, then RTF, then plain text.</summary>
    Richest = 0,
    /// <summary>Select plain text first, then HTML, then RTF.</summary>
    PlainTextFirst = 1,
    /// <summary>Select RTF, then plain text, then HTML.</summary>
    RtfFirst = 2
}

/// <summary>Policy for non-embedded resource references in email HTML.</summary>
public enum EmailRemoteResourcePolicy {
    /// <summary>Remove network resource references from the safe projection.</summary>
    Block = 0,
    /// <summary>Retain HTTP(S) references for a downstream policy-aware resolver.</summary>
    AllowByConsumerResolver = 1
}

/// <summary>Options for the dependency-isolated email body projection.</summary>
public sealed class EmailBodyProjectionOptions {
    /// <summary>Body preference.</summary>
    public EmailBodySelectionPolicy SelectionPolicy { get; set; } = EmailBodySelectionPolicy.Richest;
    /// <summary>Remote resource policy. Network access is never performed by this package.</summary>
    public EmailRemoteResourcePolicy RemoteResourcePolicy { get; set; } = EmailRemoteResourcePolicy.Block;
    /// <summary>Maximum bytes one inline resource may materialize.</summary>
    public long MaxResourceBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Optional base URI used for content-location resolution.</summary>
    public Uri? BaseUri { get; set; }
    /// <summary>Optional shared HTML policy. An untrusted profile is used when null.</summary>
    public HtmlConversionDocumentOptions? HtmlOptions { get; set; }

    internal EmailBodyProjectionOptions CloneAndValidate() {
        if (!Enum.IsDefined(typeof(EmailBodySelectionPolicy), SelectionPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(SelectionPolicy));
        }
        if (!Enum.IsDefined(typeof(EmailRemoteResourcePolicy), RemoteResourcePolicy)) {
            throw new ArgumentOutOfRangeException(nameof(RemoteResourcePolicy));
        }
        if (MaxResourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxResourceBytes));
        return new EmailBodyProjectionOptions {
            SelectionPolicy = SelectionPolicy,
            RemoteResourcePolicy = RemoteResourcePolicy,
            MaxResourceBytes = MaxResourceBytes,
            BaseUri = BaseUri,
            HtmlOptions = HtmlOptions?.Clone()
        };
    }
}

/// <summary>Operation-scoped attachment resource resolved by CID, content location, or filename.</summary>
public sealed class EmailBodyResource {
    private readonly EmailAttachment _attachment;
    private readonly long _maximumBytes;

    internal EmailBodyResource(EmailAttachment attachment, long maximumBytes) {
        _attachment = attachment;
        _maximumBytes = maximumBytes;
    }

    /// <summary>Content type declared by the artifact.</summary>
    public string ContentType => _attachment.ContentType ?? "application/octet-stream";
    /// <summary>Declared decoded length.</summary>
    public long Length => _attachment.Length;
    /// <summary>Normalized Content-ID without angle brackets.</summary>
    public string? ContentId => NormalizeContentId(_attachment.ContentId);
    /// <summary>Content-Location retained by the artifact.</summary>
    public string? ContentLocation => _attachment.ContentLocation;
    /// <summary>Safe filename retained by the artifact.</summary>
    public string? FileName => _attachment.FileName;

    /// <summary>Reads this resource within the configured bound. Each call opens a fresh operation-scoped source.</summary>
    public byte[] ReadAllBytes(CancellationToken cancellationToken = default) {
        if (Length > _maximumBytes) {
            throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                Length, _maximumBytes);
        }
        using Stream source = _attachment.OpenContentStream();
        using (var output = new MemoryStream()) {
            var buffer = new byte[64 * 1024];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                total = checked(total + read);
                if (total > _maximumBytes) {
                    throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                        total, _maximumBytes);
                }
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
    }

    /// <summary>Asynchronously reads this resource within the configured bound.</summary>
    public async Task<byte[]> ReadAllBytesAsync(CancellationToken cancellationToken = default) {
        if (Length > _maximumBytes) {
            throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                Length, _maximumBytes);
        }
        using Stream source = await _attachment.OpenContentStreamAsync(cancellationToken)
            .ConfigureAwait(false);
        using (var output = new MemoryStream()) {
            var buffer = new byte[64 * 1024];
            long total = 0;
            while (true) {
                int read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0) break;
                total = checked(total + read);
                if (total > _maximumBytes) {
                    throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                        total, _maximumBytes);
                }
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
    }

    internal static string? NormalizeContentId(string? value) => string.IsNullOrWhiteSpace(value)
        ? null
        : value!.Trim().Trim('<', '>');
}

/// <summary>One safe, reusable body/resource projection shared by transport consumers.</summary>
public sealed class EmailBodyProjectionResult {
    private readonly IReadOnlyList<EmailBodyResource> _resources;
    private readonly Uri? _baseUri;

    internal EmailBodyProjectionResult(EmailBodySourceKind sourceKind, string html,
        string? text, HtmlConversionDocument document, IReadOnlyList<EmailBodyResource> resources,
        IReadOnlyList<EmailDiagnostic> diagnostics, Uri? baseUri) {
        SourceKind = sourceKind;
        Html = html;
        Text = text;
        Document = document;
        _resources = resources;
        Diagnostics = diagnostics;
        _baseUri = baseUri;
    }

    /// <summary>Selected source body.</summary>
    public EmailBodySourceKind SourceKind { get; }
    /// <summary>Policy-normalized safe HTML.</summary>
    public string Html { get; }
    /// <summary>Original plain text when selected or available as a safe fallback.</summary>
    public string? Text { get; }
    /// <summary>Prepared canonical HTML conversion document for downstream adapters.</summary>
    public HtmlConversionDocument Document { get; }
    /// <summary>Operation-scoped embedded resources.</summary>
    public IReadOnlyList<EmailBodyResource> Resources => _resources;
    /// <summary>Stable selection, RTF fallback, and safety diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>Resolves CID, content-location, resolved absolute URI, or filename without opening content.</summary>
    public EmailBodyResource? ResolveResource(string? reference, Uri? resolvedUri = null) {
        if (string.IsNullOrWhiteSpace(reference) && resolvedUri == null) return null;
        string value = (reference ?? resolvedUri!.OriginalString).Trim();
        if (value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) {
            string id = Uri.UnescapeDataString(value.Substring(4)).Trim().Trim('<', '>');
            return _resources.FirstOrDefault(resource => string.Equals(resource.ContentId, id,
                StringComparison.OrdinalIgnoreCase));
        }
        foreach (EmailBodyResource resource in _resources) {
            if (string.Equals(resource.ContentLocation, value, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(resource.FileName, value, StringComparison.OrdinalIgnoreCase)) return resource;
            if (!string.IsNullOrWhiteSpace(resource.ContentLocation) && _baseUri != null &&
                Uri.TryCreate(_baseUri, resource.ContentLocation, out Uri? resourceUri) &&
                resolvedUri != null && string.Equals(resourceUri.AbsoluteUri, resolvedUri.AbsoluteUri,
                    StringComparison.OrdinalIgnoreCase)) return resource;
        }
        return null;
    }
}

/// <summary>Builds the canonical dependency-isolated safe email body projection.</summary>
public static class EmailBodyProjection {
    /// <summary>Projects HTML, RTF, or text and indexes embedded resources under one policy.</summary>
    public static EmailBodyProjectionResult Create(EmailDocument source,
        EmailBodyProjectionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        EmailBodyProjectionOptions effective = (options ?? new EmailBodyProjectionOptions()).CloneAndValidate();
        var diagnostics = new List<EmailDiagnostic>();
        EmailBodySourceKind sourceKind;
        string selectedHtml;
        if (effective.SelectionPolicy == EmailBodySelectionPolicy.PlainTextFirst &&
            !string.IsNullOrEmpty(source.Body.Text)) {
            sourceKind = EmailBodySourceKind.PlainText;
            selectedHtml = PlainTextHtml(source.Body.Text!);
        } else if (effective.SelectionPolicy == EmailBodySelectionPolicy.RtfFirst &&
            !string.IsNullOrWhiteSpace(source.Body.Rtf)) {
            selectedHtml = ProjectRtf(source, diagnostics, out sourceKind);
        } else if (effective.SelectionPolicy == EmailBodySelectionPolicy.RtfFirst &&
            !string.IsNullOrEmpty(source.Body.Text)) {
            sourceKind = EmailBodySourceKind.PlainText;
            selectedHtml = PlainTextHtml(source.Body.Text!);
        } else if (!string.IsNullOrWhiteSpace(source.Body.Html)) {
            sourceKind = EmailBodySourceKind.Html;
            selectedHtml = source.Body.Html!;
        } else if (!string.IsNullOrWhiteSpace(source.Body.Rtf)) {
            selectedHtml = ProjectRtf(source, diagnostics, out sourceKind);
        } else if (!string.IsNullOrEmpty(source.Body.Text)) {
            sourceKind = EmailBodySourceKind.PlainText;
            selectedHtml = PlainTextHtml(source.Body.Text!);
        } else {
            sourceKind = EmailBodySourceKind.None;
            selectedHtml = EmptyHtml();
            diagnostics.Add(new EmailDiagnostic("EMAIL_BODY_MISSING",
                "The message has no renderable HTML, RTF, or plain-text body.",
                EmailDiagnosticSeverity.Warning, "message/body"));
        }

        Uri? baseUri = effective.BaseUri ?? ResolveBaseUri(source.Body.HtmlContentLocation);
        HtmlConversionDocumentOptions htmlOptions = effective.HtmlOptions?.Clone() ??
            HtmlConversionDocumentOptions.CreateUntrustedProfile();
        htmlOptions.BaseUri ??= baseUri;
        htmlOptions.UseBodyContentsOnly = true;
        HtmlUrlPolicy resources = effective.RemoteResourcePolicy == EmailRemoteResourcePolicy.Block
            ? HtmlUrlPolicy.CreateEmbeddedResourceProfile()
            : HtmlUrlPolicy.CreateWebOnlyProfile();
        resources.AllowDataUrls = true;
        resources.AllowedUrlSchemes.Add("data");
        resources.AllowedUrlSchemes.Add("cid");
        htmlOptions.ResourceUrlPolicy = resources;
        HtmlConversionDocument sourceDocument = HtmlConversionDocument.Parse(selectedHtml, htmlOptions);
        string safeHtml = CreateSafeEmailHtml(sourceDocument);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(safeHtml, htmlOptions);
        var projectedResources = source.Attachments
            .Where(attachment => attachment.IsInline ||
                !string.IsNullOrWhiteSpace(attachment.ContentId) ||
                !string.IsNullOrWhiteSpace(attachment.ContentLocation))
            .Select(attachment => new EmailBodyResource(attachment, effective.MaxResourceBytes))
            .ToArray();
        return new EmailBodyProjectionResult(sourceKind, safeHtml, source.Body.Text,
            document, projectedResources, diagnostics.AsReadOnly(), baseUri);
    }

    private static string PlainTextHtml(string text) =>
        "<pre class=\"officeimo-email-plain\">" + WebUtility.HtmlEncode(text) + "</pre>";

    private static string EmptyHtml() =>
        "<p class=\"officeimo-email-empty\">This message has no renderable body.</p>";

    private static string ProjectRtf(EmailDocument source, ICollection<EmailDiagnostic> diagnostics,
        out EmailBodySourceKind sourceKind) {
        try {
            sourceKind = EmailBodySourceKind.Rtf;
            string html = RtfDocument.Read(source.Body.Rtf!).Document.ToHtml();
            diagnostics.Add(new EmailDiagnostic("EMAIL_BODY_RTF_PROJECTED",
                "The RTF body was projected through OfficeIMO.Html.Rtf.",
                EmailDiagnosticSeverity.Information, "message/body"));
            return html;
        } catch (Exception exception) when (exception is InvalidDataException ||
            exception is ArgumentException || exception is NotSupportedException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_BODY_RTF_UNREADABLE",
                "The RTF body could not be projected: " + exception.Message,
                EmailDiagnosticSeverity.Warning, "message/body"));
            sourceKind = string.IsNullOrEmpty(source.Body.Text)
                ? EmailBodySourceKind.None
                : EmailBodySourceKind.PlainText;
            return sourceKind == EmailBodySourceKind.PlainText
                ? PlainTextHtml(source.Body.Text!)
                : EmptyHtml();
        }
    }

    private static Uri? ResolveBaseUri(string? value) =>
        Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) ? uri : null;

    private static string CreateSafeEmailHtml(HtmlConversionDocument document) {
        AngleSharp.Html.Dom.IHtmlDocument safe = document.CreateDocumentForConversion();
        foreach (IElement element in safe.QuerySelectorAll("script,iframe,object,embed,form").ToArray()) {
            element.Remove();
        }
        foreach (IElement element in safe.All) {
            foreach (IAttr attribute in element.Attributes
                .Where(attribute => attribute.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                .ToArray()) {
                element.RemoveAttribute(attribute.Name);
            }
        }
        string styles = string.Concat(safe.Head?.QuerySelectorAll("style")
            .Select(element => element.OuterHtml) ?? Enumerable.Empty<string>());
        return styles + (safe.Body?.InnerHtml ?? safe.DocumentElement?.InnerHtml ?? string.Empty);
    }
}
