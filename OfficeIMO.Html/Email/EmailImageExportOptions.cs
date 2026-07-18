namespace OfficeIMO.Html;

/// <summary>HTML-backed email image-export options.</summary>
public sealed class EmailImageExportOptions : HtmlRenderOptions {
    /// <summary>Renders subject, sender, recipients, and date above the message body.</summary>
    public bool IncludeMessageHeaders { get; set; } = true;

    /// <summary>Uses the HTML body when available before falling back to RTF or plain text.</summary>
    public bool PreferHtmlBody { get; set; } = true;

    /// <summary>Allows MIME related attachments to satisfy content-id and content-location image references.</summary>
    public bool IncludeInlineResources { get; set; } = true;

    /// <summary>Creates an independent email options snapshot.</summary>
    public EmailImageExportOptions CloneEmail() {
        EmailImageExportOptions clone = CopyTo(new EmailImageExportOptions());
        clone.IncludeMessageHeaders = IncludeMessageHeaders;
        clone.PreferHtmlBody = PreferHtmlBody;
        clone.IncludeInlineResources = IncludeInlineResources;
        return clone;
    }

    /// <inheritdoc />
    public override HtmlRenderOptions Clone() => CloneEmail();
}
