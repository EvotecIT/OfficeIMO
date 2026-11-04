using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds sparklines to a worksheet.</summary>
/// <example>
/// <summary>Add line sparklines</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
///     ExcelSheet 'Departments' {
///         ExcelSparkline -DataRange B2:M20 -LocationRange N2:N20
///     }
/// }</code>
/// <para>Draws a sparkline for each row of monthly figures in the column next to them.</para>
/// </example>
/// <example>
/// <summary>Show weekly trends with high and low points</summary>
/// <code>New-OfficeExcel -Path .\Reports\Site-Incidents.xlsx {
///     ExcelSheet 'Incidents' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Incidents-By-Week.csv) -Name Incidents
///         ExcelSparkline -DataRange B2:M9 -LocationRange N2:N9 -Type Column -ShowHighLow
///     }
/// }</code>
/// <para>Adds column sparklines with the highest and lowest weeks marked, so each site's pattern is visible at a glance.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSparkline", DefaultParameterSetName = "Default")]
[Alias("ExcelSparkline")]
public sealed class AddOfficeExcelSparklineCommand : PSCmdlet
{

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 range containing source data (e.g., "B2:M2").</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public string DataRange { get; set; } = string.Empty;

    /// <summary>A1 range where the sparklines will be placed (e.g., "N2:N2").</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    public string LocationRange { get; set; } = string.Empty;

    /// <summary>Sparkline type (Line, Column, WinLoss).</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelSparklineType Type { get; set; } = ExcelSparklineType.Line;

    /// <summary>Show markers for each point.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowMarkers { get; set; }

    /// <summary>Show high/low points.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowHighLow { get; set; }

    /// <summary>Show first/last points.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowFirstLast { get; set; }

    /// <summary>Show negative points.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowNegative { get; set; }

    /// <summary>Show axis.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowAxis { get; set; }

    /// <summary>Series color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? SeriesColor { get; set; }

    /// <summary>Axis color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? AxisColor { get; set; }

    /// <summary>Negative point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? NegativeColor { get; set; }

    /// <summary>Markers color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? MarkerColor { get; set; }

    /// <summary>High point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? HighColor { get; set; }

    /// <summary>Low point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LowColor { get; set; }

    /// <summary>First point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FirstColor { get; set; }

    /// <summary>Last point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LastColor { get; set; }

    /// <summary>Emit the worksheet after adding sparklines.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        sheet.AddSparklines(
            dataRange: DataRange,
            locationRange: LocationRange,
            type: Type,
            displayMarkers: ShowMarkers.IsPresent,
            displayHighLow: ShowHighLow.IsPresent,
            displayFirstLast: ShowFirstLast.IsPresent,
            displayNegative: ShowNegative.IsPresent,
            displayAxis: ShowAxis.IsPresent,
            seriesColor: SeriesColor,
            axisColor: AxisColor,
            negativeColor: NegativeColor,
            markersColor: MarkerColor,
            highColor: HighColor,
            lowColor: LowColor,
            firstColor: FirstColor,
            lastColor: LastColor);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet() => ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
}
