using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets PDF diagnostics, stream statistics, feature markers, and read/rewrite blockers.</summary>
/// <example>
///   <summary>Inspect diagnostics for a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$diagnostic = Get-OfficePdfDiagnostic -Path .\Report.pdf
/// $diagnostic.StreamTypeCounts
/// $diagnostic.Findings</code>
///   <para>Returns an OfficeIMO.Pdf diagnostic report for migration and troubleshooting workflows.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfDiagnostic")]
[OutputType(typeof(PdfDiagnosticReport))]
public sealed class GetOfficePdfDiagnosticCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to analyze a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfDocument
            .Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Diagnostics());
    }
}
