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
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Shows, hides, positions and styles the chart legend.</summary>
/// <example>
/// <summary>Move the legend to the bottom</summary>
/// <code>New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
///     ExcelSheet 'Trend' {
///         ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) -Category Week -Value Opened, Closed -Type Line -Address A1 -PassThru | ExcelChartLegend -Position Bottom
///     }
/// }</code>
/// <para>Places the legend of a line chart under the plot area.</para>
/// </example>
/// <example>
/// <summary>Hide the legend of a single-series chart</summary>
/// <code>New-OfficeExcel -Path .\Reports\Availability.xlsx {
///     ExcelSheet 'Uptime' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -Title 'Uptime by site' -PassThru
///         ExcelChartLegend -Chart $chart -Hidden
///     }
/// }</code>
/// <para>Removes the legend where the chart title already names the only series.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartLegend", DefaultParameterSetName = "Default")]
[Alias("ExcelChartLegend")]
public sealed class SetOfficeExcelChartLegendCommand : PSCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Legend position.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartLegendPosition? Position { get; set; }

    /// <summary>Hide or show the column.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Overlay the legend on the chart area.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Overlay { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }

    /// <summary>Font size, interpreted as points for numbers; explicit physical units are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Text or rule colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Applies bold text. An explicit false value clears bold; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Bold { get; set; }

    /// <summary>Applies italic text. An explicit false value clears italic; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Italic { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool? hidden = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Hidden));
        if (hidden == true) Chart.HideLegend();
        else if (Position.HasValue || MyInvocation.BoundParameters.ContainsKey(nameof(Overlay)) || hidden == false)
            Chart.SetLegend(Position, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Overlay)));
        if (Size.HasValue || Color.HasValue || Font != null || MyInvocation.BoundParameters.ContainsKey(nameof(Bold)) || MyInvocation.BoundParameters.ContainsKey(nameof(Italic)))
            Chart.SetLegendTextStyle(Size, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Bold)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Italic)), Color, Font);
        if (PassThru) WriteObject(Chart);
    }

}
