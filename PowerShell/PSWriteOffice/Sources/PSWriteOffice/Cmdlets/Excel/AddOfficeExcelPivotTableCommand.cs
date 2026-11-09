using System.Collections;
using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;
/// <summary>Adds a pivot table to a worksheet.</summary>
/// <example>
/// <summary>Summarise sales in a pivot table</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelPivotTable -Range A1:F500 -Address H2 -RowField Region -DataField Amount -DataFunction Sum
///     }
/// }</code>
/// <para>Builds a pivot table of revenue by region from the order data in A1:F500.</para>
/// </example>
/// <example>
/// <summary>Pivot revenue by region and quarter</summary>
/// <code>$orders = Import-OfficeCsv -Path .\Data\Orders-2026.csv
/// New-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelTable $orders -Name Orders
///         ExcelPivotTable -Table Orders -Address J3 -Name SalesPivot -RowField Region -ColumnField Quarter -PageField Product -DataField Amount -DataFunction Sum -DataDisplayName 'Revenue' -DataNumberFormat '#,##0' -Layout Tabular -RefreshOnOpen
///     }
/// }</code>
/// <para>Writes a year of orders and pivots them by region and quarter with a product filter, naming the pivot so a slicer and conditional formats can refer to it.</para>
/// </example>
/// <example>
/// <summary>Put the pivot on its own sheet</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelPivotTable -SourceSheet Orders -Address A3 -RowField Region -ColumnField Product -DataField Amount -DataFunction Sum
///     }
/// }</code>
/// <para>Builds a summary sheet whose pivot reads all the data on the Orders sheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPivotTable", DefaultParameterSetName = "Range")]
[Alias("ExcelPivotTable")]
public sealed class AddOfficeExcelPivotTableCommand : PSCmdlet
{

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Range", Mandatory = true)]
    public string Range { get; set; } = null!;

    /// <summary>Gets or sets the source worksheet name, one-based position or native worksheet.</summary>
    [Parameter(ParameterSetName = "Range")]
    [Parameter(ParameterSetName = "SourceSheet", Mandatory = true)]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? SourceSheet { get; set; }

    /// <summary>A1 address of the target cell.</summary>
    [Parameter(Mandatory = true)]
    public string Address { get; set; } = null!;

    /// <summary>Optional pivot table name.</summary>
    [Parameter()]
    public string? Name { get; set; }

    /// <summary>Row fields (header names).</summary>
    [Parameter()]
    public string[]? RowField { get; set; }

    /// <summary>Column fields (header names).</summary>
    [Parameter()]
    public string[]? ColumnField { get; set; }

    /// <summary>Page fields (header names) used as filters.</summary>
    [Parameter()]
    public string[]? PageField { get; set; }

    /// <summary>Data fields (header names). Defaults to the last column when omitted.</summary>
    [Parameter()]
    public string[]? DataField { get; set; }

    /// <summary>Aggregation functions (Sum, Count, Average, etc.).</summary>
    [Parameter()]
    public ExcelPivotDataFunction[]? DataFunction { get; set; }

    /// <summary>Display names for data fields.</summary>
    [Parameter()]
    public string[]? DataDisplayName { get; set; }

    /// <summary>Number format codes for data fields.</summary>
    [Parameter()]
    public string[]? DataNumberFormat { get; set; }

    /// <summary>Disable row grand totals.</summary>
    [Parameter()]
    public SwitchParameter NoRowGrandTotals { get; set; }

    /// <summary>Disable column grand totals.</summary>
    [Parameter()]
    public SwitchParameter NoColumnGrandTotals { get; set; }

    /// <summary>Built-in chart style identifier.</summary>
    [Parameter()]
    [ArgumentCompleter(typeof(PSWriteOffice.Completion.OfficeExcelPivotStyleArgumentCompleter))]
    public string? Style { get; set; }

    /// <summary>Pivot layout (Compact, Outline, Tabular).</summary>
    [Parameter()]
    public ExcelPivotLayout Layout { get; set; } = ExcelPivotLayout.Compact;

    /// <summary>Show data fields on rows.</summary>
    [Parameter()]
    public SwitchParameter DataOnRows { get; set; }

    /// <summary>Show field headers.</summary>
    [Parameter()]
    public SwitchParameter ShowHeaders { get; set; }

    /// <summary>Show empty rows.</summary>
    [Parameter()]
    public SwitchParameter ShowEmptyRows { get; set; }

    /// <summary>Show empty columns.</summary>
    [Parameter()]
    public SwitchParameter ShowEmptyColumns { get; set; }

    /// <summary>Show drill indicators.</summary>
    [Parameter()]
    public SwitchParameter ShowDrill { get; set; }

    /// <summary>Allow users to drill into pivot details in Excel.</summary>
    [Parameter()]
    public SwitchParameter EnableDrill { get; set; }

    /// <summary>Optional row header caption.</summary>
    [Parameter()]
    public string? RowHeaderCaption { get; set; }

    /// <summary>Optional column header caption.</summary>
    [Parameter()]
    public string? ColumnHeaderCaption { get; set; }

    /// <summary>Optional grand total caption.</summary>
    [Parameter()]
    public string? GrandTotalCaption { get; set; }

    /// <summary>Optional missing-value caption.</summary>
    [Parameter()]
    public string? MissingCaption { get; set; }

    /// <summary>Optional error-value caption.</summary>
    [Parameter()]
    public string? ErrorCaption { get; set; }

    /// <summary>Show the data drop-down.</summary>
    [Parameter()]
    public SwitchParameter ShowDataDropDown { get; set; }

    /// <summary>Show pivot drop zones.</summary>
    [Parameter()]
    public SwitchParameter ShowDropZones { get; set; }

    /// <summary>Show pivot data tips.</summary>
    [Parameter()]
    public SwitchParameter ShowDataTips { get; set; }

    /// <summary>Show member property tips.</summary>
    [Parameter()]
    public SwitchParameter ShowMemberPropertyTips { get; set; }

    /// <summary>Controls the ordering of the pivot field list.</summary>
    [Parameter()]
    public System.ComponentModel.ListSortDirection? FieldListSort { get; set; }

    /// <summary>Use Excel custom-list sorting.</summary>
    [Parameter()]
    public SwitchParameter CustomListSort { get; set; }

    /// <summary>Field sort map, for example @{ Region = 'Ascending' }.</summary>
    [Parameter()]
    public Hashtable? FieldSort { get; set; }

    /// <summary>Field item captions to hide, for example @{ Region = @('Legacy') }.</summary>
    [Parameter()]
    public Hashtable? FieldHiddenItems { get; set; }

    /// <summary>Field item captions to keep visible, hiding other known items.</summary>
    [Parameter()]
    public Hashtable? FieldVisibleItems { get; set; }

    /// <summary>Selected page-field item captions, for example @{ Product = 'Standard' }.</summary>
    [Parameter()]
    public Hashtable? PageFieldSelection { get; set; }

    /// <summary>Fields with default subtotal disabled.</summary>
    [Parameter()]
    public string[]? FieldNoDefaultSubtotal { get; set; }

    /// <summary>Fields with subtotals shown at the top.</summary>
    [Parameter()]
    public string[]? FieldSubtotalTop { get; set; }

    /// <summary>Fields that insert blank rows after items.</summary>
    [Parameter()]
    public string[]? FieldInsertBlankRow { get; set; }

    /// <summary>Fields that insert page breaks after items.</summary>
    [Parameter()]
    public string[]? FieldInsertPageBreak { get; set; }

    /// <summary>Fields using compact field layout.</summary>
    [Parameter()]
    public string[]? FieldCompact { get; set; }

    /// <summary>Fields using outline field layout.</summary>
    [Parameter()]
    public string[]? FieldOutline { get; set; }

    /// <summary>Fields whose filter drop-downs should be hidden.</summary>
    [Parameter()]
    public string[]? FieldHideDropDowns { get; set; }

    /// <summary>Refresh the pivot cache when the workbook opens.</summary>
    [Parameter()]
    public SwitchParameter RefreshOnOpen { get; set; }

    /// <summary>Save pivot source cache records in the workbook package.</summary>
    [Parameter()]
    public SwitchParameter SaveSourceData { get; set; }

    /// <summary>Preserve pivot formatting when Excel refreshes the pivot table.</summary>
    [Parameter()]
    public SwitchParameter PreserveFormatting { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Table", Mandatory = true)]
    public string Table { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var dataFields = BuildDataFields();
        var dataOnRows = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(DataOnRows));
        var showHeaders = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowHeaders));
        var showEmptyRows = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowEmptyRows));
        var showEmptyColumns = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowEmptyColumns));
        var showDrill = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowDrill));
        var showDataDropDown = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowDataDropDown));
        var showDropZones = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowDropZones));
        var showDataTips = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowDataTips));
        var showMemberPropertyTips = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowMemberPropertyTips));
        var fieldListSortAscending = FieldListSort.HasValue ? FieldListSort == System.ComponentModel.ListSortDirection.Ascending : (bool?)null;
        var customListSort = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(CustomListSort));
        var pivotOptions = new ExcelPivotTableOptions
        {
            RefreshOnOpen = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(RefreshOnOpen)),
            SaveSourceData = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(SaveSourceData)),
            PreserveFormatting = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(PreserveFormatting)),
            EnableDrill = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(EnableDrill))
        };

        InvokeAddPivotTable(
            sheet,
            dataFields,
            Layout,
            dataOnRows,
            showHeaders,
            showEmptyRows,
            showEmptyColumns,
            showDrill,
            showDataDropDown,
            showDropZones,
            showDataTips,
            showMemberPropertyTips,
            fieldListSortAscending,
            customListSort,
            HasAnyOption(pivotOptions) ? pivotOptions : null);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private IEnumerable<ExcelPivotDataField>? BuildDataFields()
    {
        if (DataField == null || DataField.Length == 0)
        {
            if (DataFunction != null && DataFunction.Length > 0)
            {
                throw new PSArgumentException("DataFunction requires DataField to be provided.");
            }
            if (DataDisplayName != null && DataDisplayName.Length > 0)
            {
                throw new PSArgumentException("DataDisplayName requires DataField to be provided.");
            }
            if (DataNumberFormat != null && DataNumberFormat.Length > 0)
            {
                throw new PSArgumentException("DataNumberFormat requires DataField to be provided.");
            }
            return null;
        }

        var functions = DataFunction?.ToList() ?? new List<ExcelPivotDataFunction>();
        if (functions.Count > 1 && functions.Count != DataField.Length)
        {
            throw new PSArgumentException("When providing multiple DataFunction values, the count must match DataField.");
        }
        if (DataDisplayName != null && DataDisplayName.Length > 1 && DataDisplayName.Length != DataField.Length)
        {
            throw new PSArgumentException("When providing multiple DataDisplayName values, the count must match DataField.");
        }
        if (DataNumberFormat != null && DataNumberFormat.Length > 1 && DataNumberFormat.Length != DataField.Length)
        {
            throw new PSArgumentException("When providing multiple DataNumberFormat values, the count must match DataField.");
        }

        var result = new List<ExcelPivotDataField>(DataField.Length);
        for (int i = 0; i < DataField.Length; i++)
        {
            var fieldName = DataField[i];
            if (string.IsNullOrWhiteSpace(fieldName)) continue;

            var fn = functions.Count switch
            {
                0 => ExcelPivotDataFunction.Sum,
                1 => functions[0],
                _ => functions[i]
            };
            string? displayName = ResolveIndexedValue(DataDisplayName, i);
            string? numberFormat = ResolveIndexedValue(DataNumberFormat, i);
            result.Add(CreatePivotDataField(fieldName.Trim(), fn, displayName, numberFormat));
        }

        return result.Count == 0 ? null : result;
    }

    private void InvokeAddPivotTable(ExcelSheet sheet, IEnumerable<ExcelPivotDataField>? dataFields, ExcelPivotLayout layout, bool? dataOnRows, bool? showHeaders, bool? showEmptyRows, bool? showEmptyColumns, bool? showDrill, bool? showDataDropDown, bool? showDropZones, bool? showDataTips, bool? showMemberPropertyTips, bool? fieldListSortAscending, bool? customListSort, ExcelPivotTableOptions? pivotOptions)
    {
        {
            var authoredOptions = pivotOptions ?? new ExcelPivotTableOptions();
            authoredOptions.Name = Name;
            authoredOptions.RowFields = (RowField)is { } authoredRowFields1 ? System.Linq.Enumerable.ToList(authoredRowFields1) : System.Array.Empty<string>();
            authoredOptions.ColumnFields = (ColumnField)is { } authoredColumnFields2 ? System.Linq.Enumerable.ToList(authoredColumnFields2) : System.Array.Empty<string>();
            authoredOptions.PageFields = (PageField)is { } authoredPageFields3 ? System.Linq.Enumerable.ToList(authoredPageFields3) : System.Array.Empty<string>();
            authoredOptions.DataFields = (dataFields)is { } authoredDataFields4 ? System.Linq.Enumerable.ToList(authoredDataFields4) : System.Array.Empty<ExcelPivotDataField>();
            authoredOptions.ShowRowGrandTotals = !NoRowGrandTotals.IsPresent;
            authoredOptions.ShowColumnGrandTotals = !NoColumnGrandTotals.IsPresent;
            authoredOptions.PivotStyleName = Style;
            authoredOptions.Layout = layout;
            authoredOptions.DataOnRows = dataOnRows;
            authoredOptions.ShowHeaders = showHeaders;
            authoredOptions.ShowEmptyRows = showEmptyRows;
            authoredOptions.ShowEmptyColumns = showEmptyColumns;
            authoredOptions.ShowDrill = showDrill;
            authoredOptions.FieldOptions = (BuildFieldOptions())is { } authoredFieldOptions5 ? System.Linq.Enumerable.ToList(authoredFieldOptions5) : System.Array.Empty<ExcelPivotFieldOptions>();
            authoredOptions.RowHeaderCaption = RowHeaderCaption;
            authoredOptions.ColumnHeaderCaption = ColumnHeaderCaption;
            authoredOptions.GrandTotalCaption = GrandTotalCaption;
            authoredOptions.MissingCaption = MissingCaption;
            authoredOptions.ErrorCaption = ErrorCaption;
            authoredOptions.ShowDataDropDown = showDataDropDown;
            authoredOptions.ShowDropZones = showDropZones;
            authoredOptions.ShowDataTips = showDataTips;
            authoredOptions.ShowMemberPropertyTips = showMemberPropertyTips;
            authoredOptions.FieldListSortAscending = fieldListSortAscending;
            authoredOptions.CustomListSort = customListSort;
            ExcelSheet source = SourceSheet == null ? sheet : SourceSheet.Resolve(sheet.Document.Sheets, s => s.Name);
            string range = Table != null ? source.GetTableRange(Table) ?? throw new PSArgumentException("Table does not exist on the source worksheet.") : Range ?? source.UsedRange.Address;
            sheet.AddPivotTable(source, range, Address, options: authoredOptions);
        }
    }

    private static bool HasAnyOption(ExcelPivotTableOptions options)
    {
        return options.RefreshOnOpen.HasValue
            || options.SaveSourceData.HasValue
            || options.PreserveFormatting.HasValue
            || options.EnableDrill.HasValue;
    }

    private IEnumerable<ExcelPivotFieldOptions>? BuildFieldOptions()
    {
        if (!RequiresFieldOptions()) return null;

        var fields = CollectOptionFieldNames();
        return fields.Select(CreateFieldOption).ToArray();
    }

    private ExcelPivotFieldOptions CreateFieldOption(string field)
    {
        return new ExcelPivotFieldOptions(
            field,
            sortType: ResolveFieldSort(field),
            defaultSubtotal: ContainsField(FieldNoDefaultSubtotal, field) ? false : null,
            subtotalTop: ContainsField(FieldSubtotalTop, field) ? true : null,
            insertBlankRow: ContainsField(FieldInsertBlankRow, field) ? true : null,
            insertPageBreak: ContainsField(FieldInsertPageBreak, field) ? true : null,
            compact: ContainsField(FieldCompact, field) ? true : null,
            outline: ContainsField(FieldOutline, field) ? true : null,
            showDropDowns: ContainsField(FieldHideDropDowns, field) ? false : null,
            hiddenItems: ResolveMapItems(FieldHiddenItems, field),
            visibleItems: ResolveMapItems(FieldVisibleItems, field),
            selectedItem: ResolveMapScalar(PageFieldSelection, field));
    }

    private ExcelPivotFieldSort? ResolveFieldSort(string field)
    {
        string? sort = ResolveMapScalar(FieldSort, field);
        if (string.IsNullOrWhiteSpace(sort)) return null;

        if (!OpenXmlValueParser.TryParse(sort, out ExcelPivotFieldSort parsed))
        {
            throw new PSArgumentException($"Unknown field sort value '{sort}'.");
        }

        return parsed;
    }

    private SortedSet<string> CollectOptionFieldNames()
    {
        var fields = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        AddFields(fields, FieldSort?.Keys);
        AddFields(fields, FieldHiddenItems?.Keys);
        AddFields(fields, FieldVisibleItems?.Keys);
        AddFields(fields, PageFieldSelection?.Keys);
        AddFields(fields, FieldNoDefaultSubtotal);
        AddFields(fields, FieldSubtotalTop);
        AddFields(fields, FieldInsertBlankRow);
        AddFields(fields, FieldInsertPageBreak);
        AddFields(fields, FieldCompact);
        AddFields(fields, FieldOutline);
        AddFields(fields, FieldHideDropDowns);
        return fields;
    }

    private static void AddFields(SortedSet<string> fields, IEnumerable? values)
    {
        if (values == null) return;
        foreach (object? value in values)
        {
            string? text = value?.ToString();
            if (!string.IsNullOrWhiteSpace(text)) fields.Add(text!.Trim());
        }
    }

    private static string[]? ResolveMapItems(Hashtable? map, string field)
    {
        object? value = ResolveMapValue(map, field);
        if (value == null) return null;
        if (value is string text) return new[] { text };
        if (value is IEnumerable enumerable)
        {
            return enumerable.Cast<object?>()
                .Select(item => item?.ToString())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item!.Trim())
                .ToArray();
        }

        return new[] { value.ToString() ?? string.Empty };
    }

    private static string? ResolveMapScalar(Hashtable? map, string field)
    {
        object? value = ResolveMapValue(map, field);
        return value?.ToString();
    }

    private static object? ResolveMapValue(Hashtable? map, string field)
    {
        if (map == null) return null;
        foreach (DictionaryEntry entry in map)
        {
            if (entry.Key != null && string.Equals(entry.Key.ToString(), field, StringComparison.OrdinalIgnoreCase))
            {
                return entry.Value;
            }
        }

        return null;
    }

    private static bool ContainsField(string[]? fields, string field)
    {
        return fields?.Any(value => string.Equals(value, field, StringComparison.OrdinalIgnoreCase)) == true;
    }

    private bool RequiresFieldOptions()
    {
        return FieldSort?.Count > 0
            || FieldHiddenItems?.Count > 0
            || FieldVisibleItems?.Count > 0
            || PageFieldSelection?.Count > 0
            || FieldNoDefaultSubtotal?.Length > 0
            || FieldSubtotalTop?.Length > 0
            || FieldInsertBlankRow?.Length > 0
            || FieldInsertPageBreak?.Length > 0
            || FieldCompact?.Length > 0
            || FieldOutline?.Length > 0
            || FieldHideDropDowns?.Length > 0;
    }

    private static string? ResolveIndexedValue(string[]? values, int index)
    {
        if (values == null || values.Length == 0) return null;
        if (values.Length == 1) return values[0];
        return index < values.Length ? values[index] : null;
    }

    private ExcelPivotDataField CreatePivotDataField(string fieldName, ExcelPivotDataFunction function, string? displayName, string? numberFormat)
    {
        return new ExcelPivotDataField(fieldName, function, displayName, numberFormatId: null, numberFormat: numberFormat);
    }

    private ExcelSheet ResolveSheet() => ExcelSheetResolver.ResolveTarget(this, Document, Sheet);

}
