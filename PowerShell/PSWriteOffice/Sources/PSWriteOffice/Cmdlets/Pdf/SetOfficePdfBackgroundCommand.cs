using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears the generated PDF page background color.</summary>
/// <example>
///   <summary>Set a generated PDF page background.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfBackground.pdf {
///     Set-OfficePdfBackground -Color '#F8FAFC'
///     Add-OfficePdfHeading -Text 'Report on a soft background'
///     Add-OfficePdfParagraph -Text 'The background color applies to generated pages.'
/// }</code>
///   <para>Applies a page background color before adding content.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfBackground", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBackground")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfBackgroundCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Background color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Clear the generated PDF page background color.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var background = Clear.IsPresent ? null : PdfCommandUtilities.ParseColor(Color);
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument,
            page => page.Background(background));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
