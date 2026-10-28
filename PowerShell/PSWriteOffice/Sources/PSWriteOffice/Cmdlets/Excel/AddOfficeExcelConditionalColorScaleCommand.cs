using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a two-color scale conditional format to a range.</summary>
/// <example>
///   <summary>Apply a red-to-green color scale.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalColorScale -Range 'B2:B50' -StartColor '#FF0000' -EndColor '#00FF00' }</code>
///   <para>Applies a red-to-green scale to column B.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalColorScale", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalColorScale")]
public sealed class AddOfficeExcelConditionalColorScaleCommand : PSCmdlet
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

    /// <summary>A1 range to format.</summary>
    [Parameter(Position = 0)]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the target range.</summary>
    [Parameter]
    [Alias("ColumnName")]
    public string? HeaderName { get; set; }

    /// <summary>Optional table name for header-based range resolution.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Pivot table name used to resolve the target range.</summary>
    [Parameter]
    public string? PivotTableName { get; set; }

    /// <summary>Use the full pivot output range instead of the default data body range.</summary>
    [Parameter]
    public SwitchParameter PivotWholeTable { get; set; }

    /// <summary>Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.</summary>
    [Parameter]
    public int HeaderRow { get; set; }

    /// <summary>Include the header cell in the resolved range.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Start color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string StartColor { get; set; } = string.Empty;

    /// <summary>End color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string EndColor { get; set; } = string.Empty;

    /// <summary>Emit the range after applying the format.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        string targetRange = ExcelTargetRangeResolver.Resolve(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent, PivotTableName, !PivotWholeTable.IsPresent);

        sheet.AddConditionalColorScale(
            targetRange,
            OfficeColorUtilities.ToExcelArgbHex(StartColor)!,
            OfficeColorUtilities.ToExcelArgbHex(EndColor)!);

        if (PassThru.IsPresent)
        {
            WriteObject(targetRange);
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
