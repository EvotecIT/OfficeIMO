using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a named bookmark at the current generated PDF flow position.</summary>
/// <example>
///   <summary>Add a bookmark target and link to it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfBookmarks.pdf {
///     Add-OfficePdfText -Run @(
///       @{ Text = 'Jump to details'; LinkDestinationName = 'details'; Color = '#2563EB'; Underline = $true }
///     )
///     Add-OfficePdfPageBreak
///     Add-OfficePdfBookmark -Name 'details'
///     Add-OfficePdfHeading -Text 'Details' -Level 2
///   }</code>
///   <para>Creates an internal link target inside the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfBookmark", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBookmark")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfBookmarkCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Bookmark name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.Bookmark(Name));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
