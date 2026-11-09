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

/// <summary>Adds the report title and subtitle.</summary>
/// <example>
/// <summary>Add a report title</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTitle 'Operations summary'
///     }
/// }</code>
/// <para>Adds the title block at the top of a report sheet.</para>
/// </example>
/// <example>
/// <summary>Add a title with the reporting period</summary>
/// <code>$period = (Get-Date).ToString('MMMM yyyy')
/// New-OfficeExcel -Path .\Reports\Capacity-Review.xlsx {
///     ExcelReportSheet 'Capacity' {
///         ExcelReportTitle 'Capacity review' "Storage and compute, $period"
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Capacity.csv) 'Clusters'
///     }
/// }</code>
/// <para>Adds a title and a subtitle built from the current month, so a scheduled run labels itself.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTitle", DefaultParameterSetName = "Default")]
[Alias("ExcelReportTitle")]
public sealed class AddOfficeExcelReportTitleCommand : PSCmdlet {
    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Title { get; set; } = null!;

    /// <summary>Dashboard subtitle.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public string? Subtitle { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Title(Title, Subtitle);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
