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

/// <summary>Adds a legend table that explains values or colors.</summary>
/// <example>
/// <summary>Add a legend</summary>
/// <code>$legend = @(
///     [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready to ship' }
///     [pscustomobject]@{ Status = 'Review'; Meaning = 'Waiting for the owner' }
/// )
/// New-OfficeExcel -Path .\Reports\Release-Readiness.xlsx {
///     ExcelReportSheet 'Readiness' {
///         ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend
///     }
/// }</code>
/// <para>Explains the status values used in the report.</para>
/// </example>
/// <example>
/// <summary>Match legend colours to the table</summary>
/// <code>$legend = @(
///     [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready to ship' }
///     [pscustomobject]@{ Status = 'Review'; Meaning = 'Waiting for the owner' }
///     [pscustomobject]@{ Status = 'Blocked'; Meaning = 'Cannot ship until a dependency is fixed' }
/// )
/// New-OfficeExcel -Path .\Reports\Release-Readiness.xlsx {
///     ExcelReportSheet 'Readiness' {
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Release-Items.csv) 'Release 2026.11'
///         ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend -FirstColumnFillByValue @{ Ready = '#DCFCE7'; Review = '#FEF3C7'; Blocked = '#FEE2E2' }
///     }
/// }</code>
/// <para>Colours each status in the legend the same way as the rows of the readiness table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportLegend", DefaultParameterSetName = "Default")]
[Alias("ExcelReportLegend")]
public sealed class AddOfficeExcelReportLegendCommand : PSCmdlet {
    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default", Position = 0)]
    public string? Title { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public string[] Header { get; set; } = Array.Empty<string>();

    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Optional first-column fill colors keyed by first-column value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public Hashtable? FirstColumnFillByValue { get; set; }

    /// <summary>Optional header fill color.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? HeaderFillColor { get; set; }

    /// <summary>Matches names and text using ordinal case sensitivity. Omission ignores case.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().SectionLegend(Title ?? string.Empty, Header, InputObject.Select(item => ReportBlockInput.ToRow(item, Header)), ReportBlockInput.ToColorMap(FirstColumnFillByValue, CaseSensitive), HeaderFillColor);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
