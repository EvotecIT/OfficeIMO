using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds an embedded file attachment to a generated PDF document.</summary>
/// <example>
///   <summary>Embed source data in a generated PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$dataPath = '.\Examples\Documents\ServiceData.json'
/// Set-Content -Path $dataPath -Value '{ "service": "Directory", "status": "Healthy" }'
/// New-OfficePdf -Path .\Examples\Documents\PdfWithAttachment.pdf {
///     Add-OfficePdfHeading -Text 'Service report'
///     Add-OfficePdfAttachment -Path $dataPath -Name 'service-data.json' -MimeType 'application/json' -Description 'Source data used by the report.'
/// }</code>
///   <para>Embeds a supporting JSON file in the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfAttachment", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfAttachment")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfAttachmentCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>File path to embed in the generated PDF.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional embedded file name. The source file name is used when omitted.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional MIME type for the embedded file.</summary>
    [Parameter]
    public string? MimeType { get; set; }

    /// <summary>Associated-file relationship between the PDF and the embedded file.</summary>
    [Parameter]
    public PdfAssociatedFileRelationship Relationship { get; set; } = PdfAssociatedFileRelationship.Supplement;

    /// <summary>Optional human-readable attachment description.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Accepted for compatibility. The replacement document is always emitted when -Document is used.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var fileName = string.IsNullOrWhiteSpace(Name)
            ? System.IO.Path.GetFileName(inputPath)
            : Name!;
        var attachment = new PdfEmbeddedFile(fileName, File.ReadAllBytes(inputPath), MimeType, Relationship, Description);

        if (ParameterSetName == ParameterSetDocument)
        {
            var document = Document ?? throw new PSArgumentNullException(nameof(Document));
            WriteObject(document.Attachments.Add(attachment).ToDocument());
            return;
        }

        PdfCommandUtilities.ConfigureOptions(this, options => options.AddEmbeddedFile(attachment));
    }
}
