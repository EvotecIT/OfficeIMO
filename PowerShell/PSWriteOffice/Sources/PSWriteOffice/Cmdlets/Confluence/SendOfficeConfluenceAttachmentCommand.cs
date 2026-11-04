using System;
using PSWriteOffice.Foundation;
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
[Cmdlet(VerbsCommunications.Send, "OfficeConfluenceAttachment", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
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
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal local files, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

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
    public SwitchParameter NoMinorEdit { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        var paths = OfficePath.ResolveInputs(this, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath");
        if (FileName != null && paths.Count != 1) throw new PSArgumentException("A custom FileName requires one input file.");
        foreach (var resolvedPath in paths) await Upload(resolvedPath).ConfigureAwait(false);
    }

    private async Task Upload(string resolvedPath) {
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
            MinorEdit = !NoMinorEdit.IsPresent
        };
        using var client = Session.CreateClient();
        var results = await client.UploadAttachmentAsync(PageId, upload, CancelToken).ConfigureAwait(false);
        foreach (var attachment in results)
        {
            WriteObject(attachment);
        }
    }
}
