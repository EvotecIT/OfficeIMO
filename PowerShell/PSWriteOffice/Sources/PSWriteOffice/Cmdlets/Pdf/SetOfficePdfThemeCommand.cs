using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Applies an OfficeIMO.Pdf theme preset to a generated PDF document.</summary>
/// <remarks>
/// Themes provide a reusable visual baseline for generated PDFs. They are OfficeIMO.Pdf presets, exposed by PSWriteOffice as simple enum values.
/// Apply a theme near the start of a <c>New-OfficePdf</c> script block so later content inherits the intended report rhythm.
/// </remarks>
/// <example>
///   <summary>Apply the report theme.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfTheme Report
///     PdfHeading 'Service Review'
///     PdfParagraph 'The report theme defines a polished baseline.'
///   }</code>
///   <para>Uses the PDF report theme for generated content.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfTheme", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfTheme")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfThemeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Theme preset to apply.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public OfficePdfThemePreset Theme { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var theme = ResolveTheme();
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument,
            page => page.Theme(theme));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }

    private PdfTheme ResolveTheme()
    {
        return PdfThemeUtilities.ResolveTheme(Theme);
    }
}
