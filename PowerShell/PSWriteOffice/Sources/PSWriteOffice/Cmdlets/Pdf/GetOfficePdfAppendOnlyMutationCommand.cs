using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets append-only PDF mutation support and blockers for an existing PDF.</summary>
/// <example>
///   <summary>Check whether metadata can be updated incrementally.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$plan = Get-OfficePdfAppendOnlyMutation -Path .\SignedOrReviewed.pdf
/// $plan.CanAppendMetadata
/// $plan.Blockers</code>
///   <para>Returns OfficeIMO.Pdf append-only mutation support and blocker details.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfAppendOnlyMutation")]
[OutputType(typeof(PdfAppendOnlyMutationReport))]
public sealed class GetOfficePdfAppendOnlyMutationCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
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
            .AnalyzeAppendOnlyMutation());
    }
}
