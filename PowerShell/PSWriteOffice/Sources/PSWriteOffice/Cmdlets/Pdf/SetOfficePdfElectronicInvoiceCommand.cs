using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Configures Factur-X/ZUGFeRD e-invoice groundwork on a generated PDF document.</summary>
/// <example>
///   <summary>Attach CII XML as a Factur-X payload.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\Invoice.pdf {
///     Set-OfficePdfElectronicInvoice -Path .\Examples\Documents\factur-x.xml -Profile FacturX
///     Add-OfficePdfHeading -Text 'Invoice'
/// }</code>
///   <para>Embeds the XML as canonical factur-x.xml, emits matching XMP metadata, and configures PDF/A-3 groundwork.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfElectronicInvoice", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfElectronicInvoice")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfElectronicInvoiceCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Electronic-invoice options must be declared inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>CrossIndustryInvoice XML file path to embed as canonical factur-x.xml.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "XmlPath", "InvoiceXmlPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>E-invoice profile to prepare.</summary>
    [Parameter]
    public PdfComplianceProfile Profile { get; set; } = PdfComplianceProfile.FacturX;

    /// <summary>Factur-X/ZUGFeRD conformance level written to XMP metadata.</summary>
    [Parameter]
    public string ConformanceLevel { get; set; } = "EN 16931";

    /// <summary>Factur-X/ZUGFeRD schema version written to XMP metadata.</summary>
    [Parameter]
    public string Version { get; set; } = "1.0";

    /// <summary>Associated-file relationship for the embedded XML payload.</summary>
    [Parameter]
    public PdfAssociatedFileRelationship Relationship { get; set; } = PdfAssociatedFileRelationship.Data;

    /// <summary>Optional human-readable attachment description.</summary>
    [Parameter]
    public string? Description { get; set; } = "Factur-X/ZUGFeRD invoice XML";

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Profile != PdfComplianceProfile.FacturX && Profile != PdfComplianceProfile.Zugferd)
        {
            throw new PSArgumentException("Use -Profile FacturX or -Profile Zugferd for e-invoice groundwork.", nameof(Profile));
        }

        if (ParameterSetName == ParameterSetDocument)
        {
            throw new PSNotSupportedException("Electronic-invoice options must be supplied before PDF creation. Use PdfElectronicInvoice inside New-OfficePdf.");
        }

        var invoicePath = PdfCommandUtilities.ResolvePath(this, Path);
        PdfCommandUtilities.ConfigureOptions(this, options =>
            options.ConfigureElectronicInvoiceGroundworkFile(Profile, invoicePath, ConformanceLevel, Version, Relationship, Description));
    }
}
