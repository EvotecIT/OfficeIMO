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

/// <summary>Adds vertical space to the report.</summary>
/// <example>
/// <summary>Add space between blocks</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTitle 'Operations summary'
///         ExcelReportSpacer
///         ExcelReportParagraph 'Week 43, 2026'
///     }
/// }</code>
/// <para>Leaves one empty row between two report blocks.</para>
/// </example>
/// <example>
/// <summary>Separate KPIs from the detail</summary>
/// <code>New-OfficeExcel -Path .\Reports\Helpdesk-Review.xlsx {
///     ExcelReportSheet 'Review' {
///         ExcelReportKpi ([ordered]@{ Opened = 412; Closed = 398; 'Median age' = '1.6 days' })
///         ExcelReportSpacer 2
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Tickets-By-Team.csv) 'Tickets by team'
///     }
/// }</code>
/// <para>Leaves two rows between the KPI cards and the detail table to give the summary room.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSpacer", DefaultParameterSetName = "Default")]
[Alias("ExcelReportSpacer")]
public sealed class AddOfficeExcelReportSpacerCommand : PSCmdlet {
    /// <summary>Auto-fit all rows.</summary>
    [Parameter(ParameterSetName = "Default", Position = 0)]
    public int Rows { get; set; } = 1;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Spacer(Rows);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
