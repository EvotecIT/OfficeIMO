using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Inspects the current process for runtime settings that affect Excel workflows.</summary>
/// <example>
///   <summary>Check runtime culture/globalization readiness before importing spreadsheets.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = Get-OfficeExcelRuntimePreflight
/// if (-not $report.IsClean) { $report.Warnings | Write-Warning }</code>
///   <para>Returns framework, operating system, culture, and globalization-invariant mode diagnostics from OfficeIMO.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelRuntimePreflight")]
[Alias("ExcelRuntimePreflight")]
public sealed class GetOfficeExcelRuntimePreflightCommand : PSCmdlet
{
    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var report = ExcelRuntimePreflight.InspectCurrent();
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("FrameworkDescription", report.FrameworkDescription));
        record.Properties.Add(new PSNoteProperty("OSDescription", report.OSDescription));
        record.Properties.Add(new PSNoteProperty("CurrentCultureName", report.CurrentCultureName));
        record.Properties.Add(new PSNoteProperty("CurrentUICultureName", report.CurrentUICultureName));
        record.Properties.Add(new PSNoteProperty("GlobalizationInvariantMode", report.GlobalizationInvariantMode));
        record.Properties.Add(new PSNoteProperty("IsClean", report.IsClean));
        record.Properties.Add(new PSNoteProperty("Warnings", report.Warnings));
        WriteObject(record);
    }
}
