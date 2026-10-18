using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds sparklines to a worksheet.</summary>
/// <example>
///   <summary>Add a line sparkline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelSparkline -DataRange 'B2:M2' -LocationRange 'N2' }</code>
///   <para>Creates a line sparkline in N2 using B2:M2.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSparkline", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSparkline")]
public sealed class AddOfficeExcelSparklineCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range containing source data (e.g., "B2:M2").</summary>
    [Parameter(Mandatory = true)]
    public string DataRange { get; set; } = string.Empty;

    /// <summary>A1 range where the sparklines will be placed (e.g., "N2:N2").</summary>
    [Parameter(Mandatory = true)]
    public string LocationRange { get; set; } = string.Empty;

    /// <summary>Sparkline type (Line, Column, WinLoss).</summary>
    [Parameter]
    public ExcelSparklineType Type { get; set; } = ExcelSparklineType.Line;

    /// <summary>Show markers for each point.</summary>
    [Parameter]
    public SwitchParameter ShowMarkers { get; set; }

    /// <summary>Show high/low points.</summary>
    [Parameter]
    public SwitchParameter ShowHighLow { get; set; }

    /// <summary>Show first/last points.</summary>
    [Parameter]
    public SwitchParameter ShowFirstLast { get; set; }

    /// <summary>Show negative points.</summary>
    [Parameter]
    public SwitchParameter ShowNegative { get; set; }

    /// <summary>Show axis.</summary>
    [Parameter]
    public SwitchParameter ShowAxis { get; set; }

    /// <summary>Series color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? SeriesColor { get; set; }

    /// <summary>Axis color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? AxisColor { get; set; }

    /// <summary>Negative point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? NegativeColor { get; set; }

    /// <summary>Markers color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? MarkersColor { get; set; }

    /// <summary>High point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? HighColor { get; set; }

    /// <summary>Low point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LowColor { get; set; }

    /// <summary>First point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FirstColor { get; set; }

    /// <summary>Last point color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LastColor { get; set; }

    /// <summary>Emit the worksheet after adding sparklines.</summary>
    [Parameter]
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
            markersColor: MarkersColor,
            highColor: HighColor,
            lowColor: LowColor,
            firstColor: FirstColor,
            lastColor: LastColor);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
