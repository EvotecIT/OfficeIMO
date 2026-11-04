using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;
using System.Management.Automation;
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
[Alias("PdfPageBorder") ]
public sealed class SetOfficePdfPageBorderCommand : PSCmdlet
{
    private const string ParameterSetContext = "Border";

    /// <summary>Border color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Border stroke width in PDF points.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Distance from the page edge to the border path in PDF points.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [OfficeLengthTransformation]
    public OfficeLength? Inset { get; set; }

    /// <summary>Border opacity from 0 through 1.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    public double? Opacity { get; set; }

    /// <summary>Border dash style.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    public OfficeStrokeDashStyle DashStyle { get; set; } = OfficeStrokeDashStyle.Solid;

    /// <summary>Clear the generated PDF page border decoration.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Clear")]
    public SwitchParameter Clear { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PdfDslContext.Require(this).ConfigurePage( page =>
        {
            if (Clear.IsPresent) page.PageBorder(null);
            else page.PageBorder(Color, Width, Inset, Opacity, DashStyle);
        });

    }
}
