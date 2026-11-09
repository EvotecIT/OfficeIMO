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

/// <summary>Adds a row of KPI cards.</summary>
/// <example>
/// <summary>Add a row of KPI cards</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportKpi ([ordered]@{ Uptime = '99.97%'; Incidents = 4; Deployments = 37 })
///     }
/// }</code>
/// <para>Shows three headline figures as KPI cards.</para>
/// </example>
/// <example>
/// <summary>Compute KPIs from the data</summary>
/// <code>$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
/// $kpis = [ordered]@{
///     Revenue           = ($orders | Measure-Object Amount -Sum).Sum
///     Orders            = $orders.Count
///     'Average order'   = [math]::Round(($orders | Measure-Object Amount -Average).Average, 2)
///     'Active customers' = ($orders | Select-Object -ExpandProperty Customer -Unique).Count
/// }
/// New-OfficeExcel -Path .\Reports\Q3-Sales-Summary.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTitle 'Q3 sales' 'July to September 2026'
///         ExcelReportKpi $kpis -PerRow 4
///     }
/// }</code>
/// <para>Calculates the sales KPIs from the order export and shows them four to a row.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportKpi", DefaultParameterSetName = "Default")]
[Alias("ExcelReportKpi")]
public sealed class AddOfficeExcelReportKpiCommand : PSCmdlet {
    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Number of KPI cards per rendered row.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int PerRow { get; set; } = 3;

    /// <summary>Optional fill color for KPI labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LabelFillColor { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().KpiRow(ReportBlockInput.ToPairs(InputObject), PerRow, LabelFillColor);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
