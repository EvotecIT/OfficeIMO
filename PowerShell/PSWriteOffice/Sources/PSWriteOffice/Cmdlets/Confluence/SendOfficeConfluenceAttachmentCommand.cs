using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Uploads or versions a Confluence page attachment.</summary>
/// <example>
/// <summary>Upload a generated report workbook.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Send-OfficeConfluenceAttachment -Session $session -PageId 12345 -Path .\report.xlsx -Comment 'Daily refresh'</code>
/// <para>Uses Confluence's multipart attachment endpoint without automatically retrying the write.</para>
/// </example>
[Cmdlet(VerbsCommunications.Send, "OfficeConfluenceAttachment", SupportsShouldProcess = true)]
[OutputType(typeof(ConfluenceAttachment))]
public sealed class SendOfficeConfluenceAttachmentCommand : AsyncPSCmdlet
{
    /// <summary>Configured Confluence session.</summary>
    [Parameter(Mandatory = true)]
    public ConfluenceSession Session { get; set; } = null!;

    /// <summary>Page identifier.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string PageId { get; set; } = string.Empty;

    /// <summary>Local file to upload.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional attachment file name. Defaults to the local file name.</summary>
    [Parameter]
    public string? FileName { get; set; }

    /// <summary>MIME content type.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string ContentType { get; set; } = "application/octet-stream";

    /// <summary>Optional attachment version comment.</summary>
    [Parameter]
    public string? Comment { get; set; }

    /// <summary>Whether the attachment update is a minor edit.</summary>
    [Parameter]
    public bool MinorEdit { get; set; } = true;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"Attachment file '{resolvedPath}' was not found.", resolvedPath);
        }

        var uploadName = string.IsNullOrWhiteSpace(FileName) ? System.IO.Path.GetFileName(resolvedPath) : FileName!;
        if (!ShouldProcess($"Page {PageId}", $"Upload Confluence attachment '{uploadName}'"))
        {
            return;
        }

        using var content = new FileStream(
            resolvedPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            81920,
            useAsync: true);
        var upload = new ConfluenceAttachmentStreamUpload
        {
            FileName = uploadName,
            ContentType = ContentType,
            Content = content,
            Comment = Comment,
            MinorEdit = MinorEdit
        };
        using var client = Session.CreateClient();
        var results = await client.UploadAttachmentAsync(PageId, upload, CancelToken).ConfigureAwait(false);
        foreach (var attachment in results)
        {
            WriteObject(attachment);
        }
    }
}
