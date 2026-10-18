using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a data bar conditional format to a range.</summary>
/// <example>
///   <summary>Add blue data bars.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalDataBar -Range 'D2:D50' -Color '#4F81BD' }</code>
///   <para>Applies data bars to column D.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalDataBar", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalDataBar")]
public sealed class AddOfficeExcelConditionalDataBarCommand : PSCmdlet
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

    /// <summary>Bar color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string Color { get; set; } = string.Empty;

    /// <summary>Emit the range after applying the format.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        string targetRange = ExcelTargetRangeResolver.Resolve(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent, PivotTableName, !PivotWholeTable.IsPresent);

        sheet.AddConditionalDataBar(targetRange, OfficeColorUtilities.ToExcelArgbHex(Color)!);

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
