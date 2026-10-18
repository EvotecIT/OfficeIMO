using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Opens an existing PDF as an OfficeIMO.Pdf document.</summary>
/// <example>
///   <summary>Open a PDF for OfficeIMO.Pdf operations.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = Get-OfficePdf -Path .\Examples\Documents\Report.pdf
/// $pdf.Read().Text | Select-Object -First 1</code>
///   <para>Returns the OfficeIMO.Pdf document object for advanced readback or operations.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdf")]
[OutputType(typeof(PdfDocument))]
public sealed class GetOfficePdfCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open a Standard password-encrypted PDF for readback operations.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent)));
    }
}
