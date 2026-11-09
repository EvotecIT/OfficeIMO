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

/// <summary>Applies a built-in style and color preset to a chart.</summary>
/// <example>
/// <summary>Apply a chart style</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
///         ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | Set-OfficeExcelChart -Style 3
///     }
/// }</code>
/// <para>Applies built-in chart style 3 to a chart returned by ExcelChart.</para>
/// </example>
/// <example>
/// <summary>Give a dashboard chart a monochrome palette</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelSheet 'Dashboard' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -PassThru
///         Set-OfficeExcelChart -Chart $chart -Style 6 -ColorStyle 14
///     }
/// }</code>
/// <para>Uses a built-in style and colour preset for a chart that sits on a branded dashboard.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChart", DefaultParameterSetName = "Default")]
public sealed class SetOfficeExcelChartCommand : PSCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Built-in chart style identifier.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(0, int.MaxValue)]
    public int Style { get; set; } = 251;

    /// <summary>Built-in chart color-style identifier.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(0, int.MaxValue)]
    public int ColorStyle { get; set; } = 10;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Chart.ApplyStylePreset(Style, ColorStyle);
        if (PassThru) WriteObject(Chart);
    }

}
