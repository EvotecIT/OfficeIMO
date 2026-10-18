using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets PDF page size, orientation, and margins.</summary>
/// <example>
///   <summary>Set page size and margins for a generated report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfPageSetup.pdf {
///     Set-OfficePdfPageSetup -PageSize A4 -Margin 42
///     Add-OfficePdfHeading -Text 'A4 report'
///     Add-OfficePdfParagraph -Text 'The report uses custom margins.'
/// }</code>
///   <para>Applies page setup before adding generated PDF content.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfPageSetup", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageSetup")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfPageSetupCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Page size name: A4, A5, Letter, Legal, or Custom.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public string? PageSize { get; set; }

    /// <summary>Custom page width in PDF points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom page height in PDF points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use landscape orientation.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>Uniform margin in PDF points.</summary>
    [Parameter]
    public double? Margin { get; set; }

    /// <summary>Left margin in PDF points.</summary>
    [Parameter]
    public double? Left { get; set; }

    /// <summary>Top margin in PDF points.</summary>
    [Parameter]
    public double? Top { get; set; }

    /// <summary>Right margin in PDF points.</summary>
    [Parameter]
    public double? Right { get; set; }

    /// <summary>Bottom margin in PDF points.</summary>
    [Parameter]
    public double? Bottom { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var configureSize = MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
                            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
                            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
                            MyInvocation.BoundParameters.ContainsKey(nameof(Landscape));
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument, page =>
        {
            if (configureSize)
            {
                page.Size(PdfCommandUtilities.ResolvePageSize(PageSize, Width, Height, Landscape.IsPresent));
            }

            if (Margin.HasValue) page.Margin(Margin.Value);
            else if (Left.HasValue || Top.HasValue || Right.HasValue || Bottom.HasValue)
                page.Margin(Left ?? 72, Top ?? 72, Right ?? 72, Bottom ?? 72);
        });

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
