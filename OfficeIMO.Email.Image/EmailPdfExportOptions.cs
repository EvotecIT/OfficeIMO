using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Email;

/// <summary>Settings for rendering one email message as a styled PDF.</summary>
public sealed class EmailPdfExportOptions {
    private HtmlPdfExportOptions _pdf = new();

    /// <summary>Renders subject, sender, recipients, and date above the message body.</summary>
    public bool IncludeMessageHeaders { get; set; } = true;

    /// <summary>Lists file attachments with their content type and size below the message body. Inline body images are not listed.</summary>
    public bool IncludeAttachmentList { get; set; } = true;

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

    /// <summary>Page, font, resource and PDF document settings for the HTML-to-PDF renderer.</summary>
    public HtmlPdfExportOptions Pdf {
        get => _pdf;
        set => _pdf = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Creates an independent options snapshot.</summary>
    public EmailPdfExportOptions Clone() => new() {
        IncludeMessageHeaders = IncludeMessageHeaders,
        IncludeAttachmentList = IncludeAttachmentList,
        PreferHtmlBody = PreferHtmlBody,
        IncludeInlineResources = IncludeInlineResources,
        RemoteResourcePolicy = RemoteResourcePolicy,
        MaxInlineResourceCount = MaxInlineResourceCount,
        MaxTotalInlineResourceBytes = MaxTotalInlineResourceBytes,
        Pdf = Pdf.ClonePdf()
    };
}
