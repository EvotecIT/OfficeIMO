using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Turns refresh-on-open on or off for pivot caches and connections.</summary>
/// <example>
/// <summary>Refresh pivot tables on open</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
///     ExcelRefreshOnOpen -PivotTables
/// }</code>
/// <para>Makes Excel refresh every pivot cache when the workbook opens.</para>
/// </example>
/// <example>
/// <summary>Stop refreshing before archiving</summary>
/// <code>Open-OfficeExcel -Path .\Archive\Sales-2026-09.xlsx {
///     ExcelRefreshOnOpen -PivotTables -Connections -Disable -PassThru
/// }</code>
/// <para>Turns off refresh-on-open for pivot tables and connections in a month-end copy, so the archived figures stay as they were.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRefreshOnOpen", DefaultParameterSetName = "Default")]
[Alias("ExcelRefreshOnOpen")]
public sealed class SetOfficeExcelRefreshOnOpenCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Update pivot cache refresh-on-open metadata.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PivotTables { get; set; }

    /// <summary>Update workbook connection refresh-on-open metadata.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Connections { get; set; }

    /// <summary>Disable refresh-on-open instead of enabling it.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Disable { get; set; }

    /// <summary>Preserve pivot cache source data.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SavePivotSourceData { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        var result = document.SetRefreshOnOpen(!Disable, PivotTables || !Connections, Connections || !PivotTables, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(SavePivotSourceData)));
        if (PassThru) WriteObject(result);
    }

}
