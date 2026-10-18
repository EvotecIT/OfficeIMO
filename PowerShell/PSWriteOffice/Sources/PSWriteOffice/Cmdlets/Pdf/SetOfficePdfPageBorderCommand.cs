using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears the generated PDF page border decoration.</summary>
/// <example>
///   <summary>Add a subtle page border.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfPageBorder.pdf {
///     Set-OfficePdfPageBorder -Color '#CBD5E1' -Width 0.75 -Inset 24 -Opacity 0.8
///     Add-OfficePdfHeading -Text 'Bordered report'
/// }</code>
///   <para>Decorates generated pages with an OfficeIMO.Pdf page border.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfPageBorder", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageBorder")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfPageBorderCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Border color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Border stroke width in PDF points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Distance from the page edge to the border path in PDF points.</summary>
    [Parameter]
    public double? Inset { get; set; }

    /// <summary>Border opacity from 0 through 1.</summary>
    [Parameter]
    public double? Opacity { get; set; }

    /// <summary>Border dash style.</summary>
    [Parameter]
    public OfficeStrokeDashStyle DashStyle { get; set; } = OfficeStrokeDashStyle.Solid;

    /// <summary>Clear the generated PDF page border decoration.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument, page =>
        {
            if (Clear.IsPresent) page.PageBorder(null);
            else page.PageBorder(PdfCommandUtilities.ParseColor(Color), Width, Inset, Opacity, DashStyle);
        });

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
