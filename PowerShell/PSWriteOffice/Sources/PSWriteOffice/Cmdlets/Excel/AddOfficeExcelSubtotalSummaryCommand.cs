using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds grouped subtotal summary rows for a worksheet data range.</summary>
/// <example>
///   <summary>Create subtotal rows below a grouped data range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet Data { Add-OfficeExcelSubtotalSummary -GroupColumn Region -ValueColumn Sales -DataEndRow 20 }</code>
///   <para>Writes SUBTOTAL formulas below the data range and applies row outline metadata to each group.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSubtotalSummary", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSubtotalSummary", "ExcelSubtotals")]
[OutputType(typeof(ExcelSubtotalResult))]
public sealed class AddOfficeExcelSubtotalSummaryCommand : PSCmdlet
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

    /// <summary>Group column as a 1-based index, column letter, or header name.</summary>
    [Parameter(Mandatory = true)]
    [Alias("By", "GroupBy")]
    public string GroupColumn { get; set; } = string.Empty;

    /// <summary>Value columns as 1-based indexes, column letters, or header names.</summary>
    [Parameter(Mandatory = true)]
    [Alias("ValueColumns", "AggregateColumn", "AggregateColumns")]
    public string[] ValueColumn { get; set; } = Array.Empty<string>();

    /// <summary>Header row that contains source labels. Defaults to the first row of the used range.</summary>
    [Parameter]
    public int? HeaderRow { get; set; }

    /// <summary>First data row. Defaults to the row after HeaderRow.</summary>
    [Parameter]
    public int? DataStartRow { get; set; }

    /// <summary>Last data row. Defaults to the last row of the used range.</summary>
    [Parameter]
    public int? DataEndRow { get; set; }

    /// <summary>First row for the generated summary block.</summary>
    [Parameter]
    public int? SummaryStartRow { get; set; }

    /// <summary>Subtotal function.</summary>
    [Parameter]
    public ExcelSubtotalFunction Function { get; set; } = ExcelSubtotalFunction.Sum;

    /// <summary>Skip writing a summary header row.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Skip writing a grand total row.</summary>
    [Parameter]
    public SwitchParameter NoGrandTotal { get; set; }

    /// <summary>Skip applying outline metadata to detail rows.</summary>
    [Parameter]
    public SwitchParameter NoOutline { get; set; }

    /// <summary>Hide detail rows when applying outline metadata.</summary>
    [Parameter]
    public SwitchParameter HideDetailRows { get; set; }

    /// <summary>Outline level used for grouped detail rows.</summary>
    [Parameter]
    public int OutlineLevel { get; set; } = 1;

    /// <summary>Text appended to each group key in the subtotal label cell.</summary>
    [Parameter]
    public string LabelSuffix { get; set; } = " Total";

    /// <summary>Label used for the optional grand total row.</summary>
    [Parameter]
    public string GrandTotalLabel { get; set; } = "Grand Total";

    /// <summary>Emit OfficeIMO subtotal generation metadata.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var used = ResolveUsedRange(sheet);
        int headerRow = HeaderRow ?? used.FirstRow;
        int dataStartRow = DataStartRow ?? checked(headerRow + 1);
        int dataEndRow = DataEndRow ?? used.LastRow;
        int groupColumn = ResolveColumn(sheet, GroupColumn, headerRow, used.FirstColumn, used.LastColumn, nameof(GroupColumn));
        int[] valueColumns = Array.ConvertAll(ValueColumn, column => ResolveColumn(sheet, column, headerRow, used.FirstColumn, used.LastColumn, nameof(ValueColumn)));

        if (OutlineLevel < 1 || OutlineLevel > 7)
        {
            throw new PSArgumentOutOfRangeException(nameof(OutlineLevel), OutlineLevel, "Excel outline level must be between 1 and 7.");
        }

        var result = sheet.AddSubtotalSummary(new ExcelSubtotalOptions
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            DataEndRow = dataEndRow,
            GroupColumn = groupColumn,
            ValueColumns = valueColumns,
            SummaryStartRow = SummaryStartRow,
            Function = Function,
            IncludeHeader = !NoHeader.IsPresent,
            IncludeGrandTotal = !NoGrandTotal.IsPresent,
            OutlineDetailRows = !NoOutline.IsPresent,
            HideDetailRows = HideDetailRows.IsPresent,
            OutlineLevel = (byte)OutlineLevel,
            LabelSuffix = LabelSuffix,
            GrandTotalLabel = GrandTotalLabel
        });

        if (PassThru.IsPresent)
        {
            WriteObject(result);
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

    private static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) ResolveUsedRange(ExcelSheet sheet)
    {
        if (!A1.TryParseRange(sheet.UsedRangeA1, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn))
        {
            throw new PSArgumentException("Unable to resolve the worksheet used range.");
        }

        return (firstRow, firstColumn, lastRow, lastColumn);
    }

    private static int ResolveColumn(ExcelSheet sheet, string value, int headerRow, int firstColumn, int lastColumn, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException("Column value cannot be empty.", parameterName);
        }

        string trimmed = value.Trim();
        if (int.TryParse(trimmed, out int index) && index > 0)
        {
            return index;
        }

        bool lettersOnly = true;
        for (int i = 0; i < trimmed.Length; i++)
        {
            if (!char.IsLetter(trimmed[i]))
            {
                lettersOnly = false;
                break;
            }
        }

        for (int column = firstColumn; column <= lastColumn; column++)
        {
            if (sheet.TryGetCellText(headerRow, column, out string header)
                && string.Equals(header, trimmed, StringComparison.OrdinalIgnoreCase))
            {
                return column;
            }
        }

        if (lettersOnly && trimmed.Length <= 3)
        {
            int letterIndex = A1.ColumnLettersToIndex(trimmed);
            if (letterIndex > 0 && letterIndex <= A1.MaxColumns)
            {
                return letterIndex;
            }
        }

        throw new PSArgumentException($"Unable to resolve column '{value}'. Use a column index, letter, or header name.", parameterName);
    }
}
