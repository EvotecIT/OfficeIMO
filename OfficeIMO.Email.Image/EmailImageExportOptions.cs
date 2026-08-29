using OfficeIMO.Html;

namespace OfficeIMO.Email;

/// <summary>HTML-backed email image-export options.</summary>
public sealed class EmailImageExportOptions : HtmlRenderOptions {
    /// <summary>Renders subject, sender, recipients, and date above the message body.</summary>
    public bool IncludeMessageHeaders { get; set; } = true;

    /// <summary>Uses the HTML body when available before falling back to RTF or plain text.</summary>
    public bool PreferHtmlBody { get; set; } = true;

    /// <summary>Allows MIME related attachments to satisfy content-id and content-location image references.</summary>
    public bool IncludeInlineResources { get; set; } = true;

    /// <summary>Controls whether network resources remain eligible for an explicitly configured resolver.</summary>
    public EmailRemoteResourcePolicy RemoteResourcePolicy { get; set; } = EmailRemoteResourcePolicy.Block;

    /// <summary>Maximum inline resources indexed for one export.</summary>
    public int MaxInlineResourceCount { get; set; } = 128;

    /// <summary>Maximum bytes read across all inline resources for one export.</summary>
    public long MaxTotalInlineResourceBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Creates an independent email options snapshot.</summary>
    public EmailImageExportOptions CloneEmail() {
        EmailImageExportOptions clone = CopyTo(new EmailImageExportOptions());
        clone.IncludeMessageHeaders = IncludeMessageHeaders;
        clone.PreferHtmlBody = PreferHtmlBody;
        clone.IncludeInlineResources = IncludeInlineResources;
        clone.RemoteResourcePolicy = RemoteResourcePolicy;
        clone.MaxInlineResourceCount = MaxInlineResourceCount;
        clone.MaxTotalInlineResourceBytes = MaxTotalInlineResourceBytes;
        return clone;
    }

    /// <inheritdoc />
    public override HtmlRenderOptions Clone() => CloneEmail();
}
