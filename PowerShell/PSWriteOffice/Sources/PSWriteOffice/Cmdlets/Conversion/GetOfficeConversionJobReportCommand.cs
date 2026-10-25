using System;
using System.Management.Automation;
using PSWriteOffice.Services;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Reads a bounded saved conversion run report without executing its job or changing outputs.</summary>
/// <para>Requires PowerShell 7.4 or newer. Reports describe an earlier attempt; checkpoint receipts and current source/output validation determine whether artifacts can be reused.</para>
/// <example>
///   <summary>Inspect the retained attempt, then review the job before explicitly retrying.</summary>
///   <code>Get-OfficeConversionJobReport -Path ./runs/monthly.json
/// Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeConversionJobReport")]
#if !FRAMEWORK
[OutputType(typeof(OfficeConversionJobRunReport))]
#endif
public sealed class GetOfficeConversionJobReportCommand : PSCmdlet {
    /// <summary>Literal run-report JSON path. Reading does not open referenced sources, outputs or jobs.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true)]
    [Alias("LiteralPath", "FullName")]
    public string Path { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
#if FRAMEWORK
        throw new PlatformNotSupportedException("Saved conversion run reports require PowerShell 7.4 or newer.");
#else
        string path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        WriteObject(OfficeConversionJobRunReportSerializer.ReadFile(path));
#endif
    }
}
