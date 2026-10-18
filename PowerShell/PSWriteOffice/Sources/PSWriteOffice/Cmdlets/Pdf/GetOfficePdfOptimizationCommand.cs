using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets lossless PDF optimization opportunities without modifying the file.</summary>
/// <example>
///   <summary>Review optimization candidates.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = Get-OfficePdfOptimization -Path .\Report.pdf
/// $report.EstimatedSavingsBytes
/// $report.DuplicateStreams</code>
///   <para>Returns stream and duplicate-object hints before any rewrite operation is attempted.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfOptimization")]
[OutputType(typeof(PdfOptimizationReport))]
public sealed class GetOfficePdfOptimizationCommand : PSCmdlet
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
            .Optimization.Analyze());
    }
}
