using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a pivot table to a worksheet.</summary>
/// <example>
///   <summary>Create a pivot table from a report data sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Region = 'North America'; Product = 'Standard'; Sales = 125000 }
///     [pscustomobject]@{ Region = 'EMEA'; Product = 'Standard'; Sales = 98000 }
///     [pscustomobject]@{ Region = 'APAC'; Product = 'Premium'; Sales = 143000 }
/// )
/// New-OfficeExcel -Path .\SalesPivot.xlsx {
///     Add-OfficeExcelSheet -Name Data {
///         Add-OfficeExcelTable -InputObject $rows -TableName Sales -AutoFit
///         Add-OfficeExcelPivotTable -SourceRange 'A1:C4' -DestinationCell 'E2' -Name 'SalesByRegion' -RowField Region -ColumnField Product -DataField Sales -DataFunction Sum -PivotStyle PivotStyleMedium9
///     }
/// }</code>
///   <para>Writes source rows to a worksheet and creates a pivot table using the existing OfficeIMO pivot support.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPivotTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPivotTable")]
public sealed class AddOfficeExcelPivotTableCommand : PSCmdlet
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

    /// <summary>Source data range including header row (e.g., "A1:D200").</summary>
    [Parameter(Mandatory = true)]
    public string SourceRange { get; set; } = string.Empty;

    /// <summary>Top-left destination cell for the pivot table (e.g., "F2").</summary>
    [Parameter(Mandatory = true)]
    public string DestinationCell { get; set; } = string.Empty;

    /// <summary>Optional pivot table name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Row fields (header names).</summary>
    [Parameter]
    public string[]? RowField { get; set; }

    /// <summary>Column fields (header names).</summary>
    [Parameter]
    public string[]? ColumnField { get; set; }

    /// <summary>Page fields (header names) used as filters.</summary>
    [Parameter]
    public string[]? PageField { get; set; }

    /// <summary>Data fields (header names). Defaults to the last column when omitted.</summary>
    [Parameter]
    public string[]? DataField { get; set; }

    /// <summary>Aggregation functions (Sum, Count, Average, etc.).</summary>
    [Parameter]
    public ExcelPivotDataFunction[]? DataFunction { get; set; }

    /// <summary>Display names for data fields.</summary>
    [Parameter]
    public string[]? DataDisplayName { get; set; }

    /// <summary>Number format codes for data fields.</summary>
    [Parameter]
    public string[]? DataNumberFormat { get; set; }

    /// <summary>Disable row grand totals.</summary>
    [Parameter]
    public SwitchParameter NoRowGrandTotals { get; set; }

    /// <summary>Disable column grand totals.</summary>
    [Parameter]
    public SwitchParameter NoColumnGrandTotals { get; set; }

    /// <summary>Optional pivot table style name.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeExcelPivotStyleArgumentCompleter))]
    public string? PivotStyle { get; set; }

    /// <summary>Pivot layout (Compact, Outline, Tabular).</summary>
    [Parameter]
    public ExcelPivotLayout Layout { get; set; } = ExcelPivotLayout.Compact;

    /// <summary>Show data fields on rows.</summary>
    [Parameter]
    public SwitchParameter DataOnRows { get; set; }

    /// <summary>Show data fields on columns.</summary>
    [Parameter]
    public SwitchParameter DataOnColumns { get; set; }

    /// <summary>Show field headers.</summary>
    [Parameter]
    public SwitchParameter ShowHeaders { get; set; }

    /// <summary>Hide field headers.</summary>
    [Parameter]
    public SwitchParameter HideHeaders { get; set; }

    /// <summary>Show empty rows.</summary>
    [Parameter]
    public SwitchParameter ShowEmptyRows { get; set; }

    /// <summary>Hide empty rows.</summary>
    [Parameter]
    public SwitchParameter HideEmptyRows { get; set; }

    /// <summary>Show empty columns.</summary>
    [Parameter]
    public SwitchParameter ShowEmptyColumns { get; set; }

    /// <summary>Hide empty columns.</summary>
    [Parameter]
    public SwitchParameter HideEmptyColumns { get; set; }

    /// <summary>Show drill indicators.</summary>
    [Parameter]
    public SwitchParameter ShowDrill { get; set; }

    /// <summary>Hide drill indicators.</summary>
    [Parameter]
    public SwitchParameter HideDrill { get; set; }

    /// <summary>Optional row header caption.</summary>
    [Parameter]
    public string? RowHeaderCaption { get; set; }

    /// <summary>Optional column header caption.</summary>
    [Parameter]
    public string? ColumnHeaderCaption { get; set; }

    /// <summary>Optional grand total caption.</summary>
    [Parameter]
    public string? GrandTotalCaption { get; set; }

    /// <summary>Optional missing-value caption.</summary>
    [Parameter]
    public string? MissingCaption { get; set; }

    /// <summary>Optional error-value caption.</summary>
    [Parameter]
    public string? ErrorCaption { get; set; }

    /// <summary>Show the data drop-down.</summary>
    [Parameter]
    public SwitchParameter ShowDataDropDown { get; set; }

    /// <summary>Hide the data drop-down.</summary>
    [Parameter]
    public SwitchParameter HideDataDropDown { get; set; }

    /// <summary>Show pivot drop zones.</summary>
    [Parameter]
    public SwitchParameter ShowDropZones { get; set; }

    /// <summary>Hide pivot drop zones.</summary>
    [Parameter]
    public SwitchParameter HideDropZones { get; set; }

    /// <summary>Show pivot data tips.</summary>
    [Parameter]
    public SwitchParameter ShowDataTips { get; set; }

    /// <summary>Hide pivot data tips.</summary>
    [Parameter]
    public SwitchParameter HideDataTips { get; set; }

    /// <summary>Show member property tips.</summary>
    [Parameter]
    public SwitchParameter ShowMemberPropertyTips { get; set; }

    /// <summary>Hide member property tips.</summary>
    [Parameter]
    public SwitchParameter HideMemberPropertyTips { get; set; }

    /// <summary>Sort pivot field list ascending.</summary>
    [Parameter]
    public SwitchParameter FieldListSortAscending { get; set; }

    /// <summary>Sort pivot field list descending.</summary>
    [Parameter]
    public SwitchParameter FieldListSortDescending { get; set; }

    /// <summary>Use Excel custom-list sorting.</summary>
    [Parameter]
    public SwitchParameter CustomListSort { get; set; }

    /// <summary>Disable Excel custom-list sorting.</summary>
    [Parameter]
    public SwitchParameter NoCustomListSort { get; set; }

    /// <summary>Field sort map, for example @{ Region = 'Ascending' }.</summary>
    [Parameter]
    public Hashtable? FieldSort { get; set; }

    /// <summary>Field item captions to hide, for example @{ Region = @('Legacy') }.</summary>
    [Parameter]
    public Hashtable? FieldHiddenItems { get; set; }

    /// <summary>Field item captions to keep visible, hiding other known items.</summary>
    [Parameter]
    public Hashtable? FieldVisibleItems { get; set; }

    /// <summary>Selected page-field item captions, for example @{ Product = 'Standard' }.</summary>
    [Parameter]
    public Hashtable? PageFieldSelection { get; set; }

    /// <summary>Fields with default subtotal disabled.</summary>
    [Parameter]
    public string[]? FieldNoDefaultSubtotal { get; set; }

    /// <summary>Fields with subtotals shown at the top.</summary>
    [Parameter]
    public string[]? FieldSubtotalTop { get; set; }

    /// <summary>Fields that insert blank rows after items.</summary>
    [Parameter]
    public string[]? FieldInsertBlankRow { get; set; }

    /// <summary>Fields that insert page breaks after items.</summary>
    [Parameter]
    public string[]? FieldInsertPageBreak { get; set; }

    /// <summary>Fields using compact field layout.</summary>
    [Parameter]
    public string[]? FieldCompact { get; set; }

    /// <summary>Fields using outline field layout.</summary>
    [Parameter]
    public string[]? FieldOutline { get; set; }

    /// <summary>Fields whose filter drop-downs should be hidden.</summary>
    [Parameter]
    public string[]? FieldHideDropDowns { get; set; }

    /// <summary>Refresh the pivot cache when the workbook opens.</summary>
    [Parameter]
    public SwitchParameter RefreshOnOpen { get; set; }

    /// <summary>Do not refresh the pivot cache when the workbook opens.</summary>
    [Parameter]
    public SwitchParameter NoRefreshOnOpen { get; set; }

    /// <summary>Save pivot source cache records in the workbook package.</summary>
    [Parameter]
    public SwitchParameter SaveSourceData { get; set; }

    /// <summary>Do not save pivot source cache records in the workbook package.</summary>
    [Parameter]
    public SwitchParameter NoSaveSourceData { get; set; }

    /// <summary>Preserve pivot formatting when Excel refreshes the pivot table.</summary>
    [Parameter]
    public SwitchParameter PreserveFormatting { get; set; }

    /// <summary>Do not preserve pivot formatting when Excel refreshes the pivot table.</summary>
    [Parameter]
    public SwitchParameter NoPreserveFormatting { get; set; }

    /// <summary>Allow users to drill into pivot details in Excel.</summary>
    [Parameter]
    public SwitchParameter EnableDrill { get; set; }

    /// <summary>Disable pivot detail drill interaction in Excel.</summary>
    [Parameter]
    public SwitchParameter DisableDrill { get; set; }

    /// <summary>Emit the worksheet after creating the pivot table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var dataFields = BuildDataFields();
        var dataOnRows = ResolveToggle(DataOnRows, DataOnColumns, "DataOnRows/DataOnColumns");
        var showHeaders = ResolveToggle(ShowHeaders, HideHeaders, "ShowHeaders/HideHeaders");
        var showEmptyRows = ResolveToggle(ShowEmptyRows, HideEmptyRows, "ShowEmptyRows/HideEmptyRows");
        var showEmptyColumns = ResolveToggle(ShowEmptyColumns, HideEmptyColumns, "ShowEmptyColumns/HideEmptyColumns");
        var showDrill = ResolveToggle(ShowDrill, HideDrill, "ShowDrill/HideDrill");
        var showDataDropDown = ResolveToggle(ShowDataDropDown, HideDataDropDown, "ShowDataDropDown/HideDataDropDown");
        var showDropZones = ResolveToggle(ShowDropZones, HideDropZones, "ShowDropZones/HideDropZones");
        var showDataTips = ResolveToggle(ShowDataTips, HideDataTips, "ShowDataTips/HideDataTips");
        var showMemberPropertyTips = ResolveToggle(ShowMemberPropertyTips, HideMemberPropertyTips, "ShowMemberPropertyTips/HideMemberPropertyTips");
        var fieldListSortAscending = ResolveToggle(FieldListSortAscending, FieldListSortDescending, "FieldListSortAscending/FieldListSortDescending");
        var customListSort = ResolveToggle(CustomListSort, NoCustomListSort, "CustomListSort/NoCustomListSort");
        var pivotOptions = new ExcelPivotTableOptions
        {
            RefreshOnOpen = ResolveToggle(RefreshOnOpen, NoRefreshOnOpen, "RefreshOnOpen/NoRefreshOnOpen"),
            SaveSourceData = ResolveToggle(SaveSourceData, NoSaveSourceData, "SaveSourceData/NoSaveSourceData"),
            PreserveFormatting = ResolveToggle(PreserveFormatting, NoPreserveFormatting, "PreserveFormatting/NoPreserveFormatting"),
            EnableDrill = ResolveToggle(EnableDrill, DisableDrill, "EnableDrill/DisableDrill")
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

    private void InvokeAddPivotTable(
        ExcelSheet sheet,
        IEnumerable<ExcelPivotDataField>? dataFields,
        ExcelPivotLayout layout,
        bool? dataOnRows,
        bool? showHeaders,
        bool? showEmptyRows,
        bool? showEmptyColumns,
        bool? showDrill,
        bool? showDataDropDown,
        bool? showDropZones,
        bool? showDataTips,
        bool? showMemberPropertyTips,
        bool? fieldListSortAscending,
        bool? customListSort,
        ExcelPivotTableOptions? pivotOptions)
    {
        sheet.AddPivotTable(
            SourceRange,
            DestinationCell,
            Name,
            RowField,
            ColumnField,
            PageField,
            dataFields,
            showRowGrandTotals: !NoRowGrandTotals.IsPresent,
            showColumnGrandTotals: !NoColumnGrandTotals.IsPresent,
            pivotStyleName: PivotStyle,
            layout: layout,
            dataOnRows: dataOnRows,
            showHeaders: showHeaders,
            showEmptyRows: showEmptyRows,
            showEmptyColumns: showEmptyColumns,
            showDrill: showDrill,
            fieldOptions: BuildFieldOptions(),
            rowHeaderCaption: RowHeaderCaption,
            columnHeaderCaption: ColumnHeaderCaption,
            grandTotalCaption: GrandTotalCaption,
            missingCaption: MissingCaption,
            errorCaption: ErrorCaption,
            showDataDropDown: showDataDropDown,
            showDropZones: showDropZones,
            showDataTips: showDataTips,
            showMemberPropertyTips: showMemberPropertyTips,
            fieldListSortAscending: fieldListSortAscending,
            customListSort: customListSort,
            options: pivotOptions);
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

    private static bool? ResolveToggle(SwitchParameter enable, SwitchParameter disable, string name)
    {
        if (enable.IsPresent && disable.IsPresent)
        {
            throw new PSArgumentException($"Cannot set both {name}.");
        }

        if (enable.IsPresent) return true;
        if (disable.IsPresent) return false;
        return null;
    }
}
