using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a highlighted info, success, warning or error callout.</summary>
/// <example>
/// <summary>Add an information callout</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportCallout Info 'Data source' 'Monitoring export from 1 November, 06:00.'
///     }
/// }</code>
/// <para>Highlights a note for readers in an information callout.</para>
/// </example>
/// <example>
/// <summary>Report results with matching callouts</summary>
/// <code>$failed = Import-OfficeCsv -Path .\Data\Backup-Results.csv | Where-Object Result -ne 'Success'
/// New-OfficeExcel -Path .\Reports\Backup-Check.xlsx {
///     ExcelReportSheet 'Backups' {
///         ExcelReportTitle 'Nightly backup check' (Get-Date).ToString('d MMMM yyyy')
///         if ($failed) {
///             ExcelReportCallout Error 'Backups failed' "$(@($failed).Count) jobs did not complete. Details are below."
///             ExcelReportTable $failed 'Failed jobs'
///         } else {
///             ExcelReportCallout Success 'All backups completed' 'Every scheduled job finished successfully.'
///         }
///     }
/// }</code>
/// <para>Adds a success or error callout depending on whether the backup checks passed, so the outcome is the first thing a reader sees.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportCallout", DefaultParameterSetName = "Default")]
[Alias("ExcelReportCallout")]
public sealed class AddOfficeExcelReportCalloutCommand : PSCmdlet {
    /// <summary>Metadata kind to add.</summary>
    [Parameter(ParameterSetName = "Default", Position = 0)]
    public OfficeExcelReportCalloutKind Kind { get; set; } = OfficeExcelReportCalloutKind.Info;

    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 1)]
    public string Title { get; set; } = null!;

    /// <summary>Comment text.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 2)]
    public string Text { get; set; } = null!;

    /// <summary>Number of worksheet columns occupied by the report block.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int Span { get; set; } = 8;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        var kind = Kind switch { OfficeExcelReportCalloutKind.Success => OfficeCalloutKind.Success, OfficeExcelReportCalloutKind.Warning => OfficeCalloutKind.Warning, OfficeExcelReportCalloutKind.Error or OfficeExcelReportCalloutKind.Critical => OfficeCalloutKind.Danger, _ => OfficeCalloutKind.Info };
        context.RequireComposer().Callout(kind, Title, Text, Span);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
