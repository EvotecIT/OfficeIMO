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

/// <summary>Adds a section heading to the report.</summary>
/// <example>
/// <summary>Add a section heading</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportHeading 'Incidents'
///     }
/// }</code>
/// <para>Starts a new section of the report.</para>
/// </example>
/// <example>
/// <summary>Split a report into sections</summary>
/// <code>New-OfficeExcel -Path .\Reports\Security-Review.xlsx {
///     ExcelReportSheet 'Review' {
///         ExcelReportTitle 'Security review' 'Q3 2026'
///         ExcelReportHeading 'Privileged accounts'
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Privileged-Accounts.csv)
///         ExcelReportHeading 'Firewall changes'
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Firewall-Changes.csv)
///     }
/// }</code>
/// <para>Gives each part of a security report its own heading and table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportHeading", DefaultParameterSetName = "Default")]
[Alias("ExcelReportHeading")]
public sealed class AddOfficeExcelReportHeadingCommand : PSCmdlet {
    /// <summary>Comment text.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Text { get; set; } = null!;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Section(Text);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
