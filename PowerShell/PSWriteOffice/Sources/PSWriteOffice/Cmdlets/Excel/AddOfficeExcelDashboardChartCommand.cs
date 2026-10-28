using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a dashboard-ready chart using an OfficeIMO chart preset.</summary>
/// <example>
///   <summary>Add a compact comparison chart.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Dashboard' { Add-OfficeExcelDashboardChart -Range A1:B12 -Preset CompactComparison -Row 1 -Column 5 -Title 'Revenue' }</code>
///   <para>Creates a styled chart from the range using reusable OfficeIMO dashboard chart defaults.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelDashboardChart", DefaultParameterSetName = ParameterSetContextRange)]
[Alias("ExcelDashboardChart")]
public sealed class AddOfficeExcelDashboardChartCommand : PSCmdlet
{
    private const string ParameterSetContextRange = "ContextRange";
    private const string ParameterSetContextTable = "ContextTable";
    private const string ParameterSetDocumentRange = "DocumentRange";
    private const string ParameterSetDocumentTable = "DocumentTable";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRange)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentTable)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range containing chart data.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextRange)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentRange)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Table name containing chart data.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextTable)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentTable)]
    public string TableName { get; set; } = string.Empty;

    /// <summary>Dashboard chart preset.</summary>
    [Parameter]
    public ExcelDashboardChartPreset Preset { get; set; } = ExcelDashboardChartPreset.Comparison;

    /// <summary>Top-left row (1-based) where the chart should be placed.</summary>
    [Parameter(Mandatory = true)]
    public int Row { get; set; }

    /// <summary>Top-left column (1-based) where the chart should be placed.</summary>
    [Parameter(Mandatory = true)]
    public int Column { get; set; }

    /// <summary>Optional chart type override.</summary>
    [Parameter]
    public ExcelChartType? ChartType { get; set; }

    /// <summary>Chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Whether the range includes headers.</summary>
    [Parameter(ParameterSetName = ParameterSetContextRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    public bool HasHeaders { get; set; } = true;

    /// <summary>Include cached data in the chart for portability.</summary>
    [Parameter]
    public bool IncludeCachedData { get; set; } = true;

    /// <summary>Optional chart width in pixels.</summary>
    [Parameter]
    public int? WidthPixels { get; set; }

    /// <summary>Optional chart height in pixels.</summary>
    [Parameter]
    public int? HeightPixels { get; set; }

    /// <summary>Optional chart style id override.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? StyleId { get; set; }

    /// <summary>Optional chart color style id override.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? ColorStyleId { get; set; }

    /// <summary>Emit the created chart.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Row <= 0 || Column <= 0)
        {
            throw new PSArgumentException("Row and Column must be 1 or greater.");
        }

        var sheet = ResolveSheet();
        var options = new ExcelDashboardChartOptions
        {
            Preset = Preset,
            Row = Row,
            Column = Column,
            ChartType = ChartType,
            Title = Title,
            HasHeaders = HasHeaders,
            IncludeCachedData = IncludeCachedData,
            WidthPixels = WidthPixels,
            HeightPixels = HeightPixels,
            StyleId = StyleId,
            ColorStyleId = ColorStyleId
        };

        if (ParameterSetName == ParameterSetContextTable || ParameterSetName == ParameterSetDocumentTable)
        {
            var tableRange = TryGetContextTableRange(sheet);
            if (!string.IsNullOrWhiteSpace(tableRange))
            {
                options.Range = tableRange;
                options.HasHeaders = true;
            }
            else
            {
                options.TableName = TableName;
            }
        }
        else
        {
            options.Range = Range;
        }

        var chart = sheet.AddDashboardChart(options);
        if (PassThru.IsPresent)
        {
            WriteObject(chart);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentRange || ParameterSetName == ParameterSetDocumentTable)
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

    private string? TryGetContextTableRange(ExcelSheet sheet)
    {
        if (ParameterSetName != ParameterSetContextTable)
        {
            return null;
        }

        try
        {
            var context = ExcelDslContext.Require(this);
            return context.TryGetTableRange(sheet, TableName, out var range) ? range : null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }
}
