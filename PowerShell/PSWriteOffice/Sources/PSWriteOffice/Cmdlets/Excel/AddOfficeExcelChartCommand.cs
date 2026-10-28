using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a chart to the current worksheet using a range or table.</summary>
/// <example>
///   <summary>Add a chart from a range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelChart -Range 'A1:D10' -Row 2 -Column 6 -Type Line -Title 'Trend' }</code>
///   <para>Creates a line chart from A1:D10 and places it at F2.</para>
/// </example>
/// <example>
///   <summary>Add a chart from a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelChart -TableName 'Sales' -Row 2 -Column 6 -Type ColumnClustered }</code>
///   <para>Creates a chart from the Sales table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelChart", DefaultParameterSetName = ParameterSetContextRange)]
[Alias("ExcelChart")]
public sealed class AddOfficeExcelChartCommand : PSCmdlet
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

    /// <summary>Top-left row (1-based) where the chart should be placed.</summary>
    [Parameter(Mandatory = true)]
    public int Row { get; set; }

    /// <summary>Top-left column (1-based) where the chart should be placed.</summary>
    [Parameter(Mandatory = true)]
    public int Column { get; set; }

    /// <summary>Chart width in pixels.</summary>
    [Parameter]
    public int WidthPixels { get; set; } = 640;

    /// <summary>Chart height in pixels.</summary>
    [Parameter]
    public int HeightPixels { get; set; } = 360;

    /// <summary>Chart type.</summary>
    [Parameter]
    public ExcelChartType Type { get; set; } = ExcelChartType.ColumnClustered;

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

        ExcelChart chart;
        if (ParameterSetName == ParameterSetContextTable || ParameterSetName == ParameterSetDocumentTable)
        {
            var tableRange = TryGetContextTableRange(sheet) ?? sheet.GetTableRange(TableName);
            if (string.IsNullOrWhiteSpace(tableRange))
            {
                throw new InvalidOperationException($"Table '{TableName}' was not found on sheet '{sheet.Name}'.");
            }

            chart = sheet.AddChartFromRange(tableRange!, Row, Column, WidthPixels, HeightPixels, Type, hasHeaders: true, Title, IncludeCachedData);
        }
        else
        {
            chart = sheet.AddChartFromRange(Range, Row, Column, WidthPixels, HeightPixels, Type, HasHeaders, Title, IncludeCachedData);
        }

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
