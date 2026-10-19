using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets generated PDF compliance profile and readiness groundwork.</summary>
/// <example>
///   <summary>Configure PDF/A groundwork before saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$path = '.\Examples\Documents\PdfCompliance.pdf'
/// New-OfficePdf -Path $path {
///     Set-OfficePdfCompliance -Profile PdfA3B -Groundwork -Language 'en-US'
///     Add-OfficePdfHeading -Text 'Compliance-ready report'
/// }
/// Get-OfficePdfCompliance -Path $path -Profile PdfA3B</code>
///   <para>Applies OfficeIMO.Pdf compliance groundwork during composition, saves the PDF, then inspects the saved file.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfCompliance", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfCompliance")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfComplianceCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Compliance options must be declared inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Requested generated PDF compliance profile.</summary>
    [Parameter(Mandatory = true)]
    public PdfComplianceProfile Profile { get; set; }

    /// <summary>Configure common PDF/A or PDF/UA groundwork for the selected profile.</summary>
    [Parameter]
    public SwitchParameter Groundwork { get; set; }

    /// <summary>Catalog language used by compliance groundwork.</summary>
    [Parameter]
    public string Language { get; set; } = "en-US";

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            throw new PSNotSupportedException("Document-wide compliance options must be supplied before PDF creation. Use PdfCompliance inside New-OfficePdf.");
        }

        PdfCommandUtilities.ConfigureOptions(this, options =>
        {
            if (Groundwork.IsPresent) ApplyGroundwork(options);
            else options.RequireCompliance(Profile);
        });
    }

    private void ApplyGroundwork(PdfOptions options)
    {
        switch (Profile)
        {
            case PdfComplianceProfile.PdfUa1:
            case PdfComplianceProfile.PdfUa2:
                options.ConfigurePdfUaGroundwork(Profile, Language);
                break;
            case PdfComplianceProfile.PdfA2B:
            case PdfComplianceProfile.PdfA2U:
            case PdfComplianceProfile.PdfA2A:
            case PdfComplianceProfile.PdfA3B:
            case PdfComplianceProfile.PdfA3U:
            case PdfComplianceProfile.PdfA3A:
            case PdfComplianceProfile.PdfA4:
            case PdfComplianceProfile.PdfA4E:
            case PdfComplianceProfile.PdfA4F:
                options.ConfigurePdfAGroundwork(Profile, Language);
                break;
            default:
                WriteWarning("Groundwork is currently available for PDF/A and PDF/UA profiles.");
                break;
        }
    }
}
