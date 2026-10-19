using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes tabular data to the current worksheet and formats it as an Excel table.</summary>
/// <para>Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.</para>
/// <example>
///   <summary>Insert a table starting at A1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
///   ExcelSheet 'Data' { Add-OfficeExcelTable -InputObject $data -TableName 'Sales' }</code>
///   <para>Writes two rows and formats them as a styled Excel table.</para>
/// </example>
/// <example>
///   <summary>Add a table to an explicit worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeExcelTable -Worksheet $sheet -InputObject $rows -TableName 'Sales' -AutoFit</code>
///   <para>Writes the rows into a live workbook without requiring an active DSL scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTable")]
[Alias("ExcelTable")]
public sealed class AddOfficeExcelTableCommand : PSCmdlet
{
    private readonly List<object?> _items = new();

    /// <summary>Worksheet that will receive the table outside a DSL context.</summary>
    [Parameter]
    [Alias("SheetObject")]
    public ExcelSheet? Worksheet { get; set; }

    /// <summary>Workbook that will receive the table outside a DSL context.</summary>
    [Parameter]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Source objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences to convert into table rows.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("Data", "DataTable")]
    public object? InputObject { get; set; }

    /// <summary>Starting row for the data (1-based).</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Starting column for the data (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Skip writing headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Text used between items when a cell contains a collection.</summary>
    [Parameter]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Text used between entries when a cell contains a dictionary.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryEntrySeparator { get; set; } = "; ";

    /// <summary>Text used between a dictionary key and value.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    /// <summary>Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Name to assign to the table.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Emphasize the first table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasize the last table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Disable alternating row stripes for the created table.</summary>
    [Parameter]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Enable alternating column stripes for the created table.</summary>
    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Disable AutoFilter dropdowns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Auto-fit the table columns after insertion.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Return the created range string.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        var sheet = ResolveSheet();

        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var projectedRows = TableViewProjection.Project(rows, View);
        var normalizerOptions = PowerShellObjectNormalizerOptions.ForTable(
            CollectionSeparator,
            DictionaryEntrySeparator,
            DictionaryKeyValueSeparator,
            MaxCollectionItems,
            MaxNestingDepth);
        var table = ExcelTabularInputService.ToDataTable(
            projectedRows,
            TableName,
            normalizerOptions: normalizerOptions);
        if (table.Columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer columns from the supplied data.");
        }

        if (StartRow < 1 || StartColumn < 1)
        {
            throw new ArgumentOutOfRangeException("StartRow/StartColumn must be 1 or greater.");
        }

        var resolvedTableName = ResolveTableName(table);
        var range = sheet.InsertDataTableAsTable(
            table,
            startRow: StartRow,
            startColumn: StartColumn,
            includeHeaders: !NoHeader.IsPresent,
            tableName: resolvedTableName,
            style: TableStyle,
            includeAutoFilter: !NoAutoFilter.IsPresent);
        ExcelTableStyleOptionService.Apply(
            sheet,
            range,
            TableStyle,
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
        var context = ExcelDslContext.Current;
        context?.RegisterTableRange(sheet, resolvedTableName, range);

        if (AutoFit.IsPresent)
        {
            var columnIndexes = Enumerable.Range(StartColumn, table.Columns.Count);
            sheet.AutoFitColumnsFor(columnIndexes);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(range);
        }
    }

    private string? ResolveTableName(System.Data.DataTable table)
    {
        if (!string.IsNullOrWhiteSpace(TableName))
        {
            return TableName;
        }

        return string.IsNullOrWhiteSpace(table.TableName) ? null : table.TableName;
    }

    private ExcelSheet ResolveSheet()
    {
        if (Worksheet != null && Document != null)
        {
            throw new PSArgumentException("Use either -Worksheet or -Document, not both.");
        }

        if (Worksheet != null)
        {
            return Worksheet;
        }

        if (Document != null)
        {
            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }
}
