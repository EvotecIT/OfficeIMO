using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a titled, styled table to the report.</summary>
/// <example>
/// <summary>Add a titled table</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTable (Get-Service | Where-Object Status -eq 'Stopped') 'Stopped services'
///     }
/// }</code>
/// <para>Adds a styled table with a title to the report.</para>
/// </example>
/// <example>
/// <summary>Choose the columns to show</summary>
/// <code>$assets = Import-OfficeCsv -Path .\Data\Assets.csv
/// New-OfficeExcel -Path .\Reports\Asset-Review.xlsx {
///     ExcelReportSheet 'Assets' {
///         ExcelReportTitle 'Laptops due for replacement'
///         ExcelReportTable ($assets | Where-Object AgeYears -ge 4) 'Due this year' -Property AssetTag, User, Model, Purchased, AgeYears -Name DueLaptops -ColumnFormat @{ Purchased = 'Date'; AgeYears = '0.0' }
///     }
/// }</code>
/// <para>Shows only the columns readers need from a wide asset export and leaves the internal identifiers out.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTable", DefaultParameterSetName = "Default")]
[Alias("ExcelReportTable")]
public sealed class AddOfficeExcelReportTableCommand : PSCmdlet {
    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public string? Title { get; set; }

    /// <summary>Named range name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Name { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Properties to include, in the requested column order.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string[]? Property { get; set; }

    /// <summary>Exclude specific properties from exported objects.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string[]? ExcludeProperty { get; set; }

    /// <summary>Maps column headers to number-format presets, format codes or per-column option hashtables.</summary>
    [Parameter(ParameterSetName = "Default")]
    public Hashtable? ColumnFormat { get; set; }

    /// <summary>Text inserted between collection values flattened into one cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Nested-value projection policy; accepts a native options object or hashtable.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeOptionsTransformation(typeof(OfficeObjectFlattenOptions))]
    public OfficeObjectFlattenOptions? FlattenOptions { get; set; }

    /// <summary>Emphasizes the first table column; explicit false removes the emphasis.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasizes the last table column; explicit false removes the emphasis.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Disables alternating row bands. An explicit false value enables them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Enables alternating column bands; explicit false clears them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Omits table filter buttons.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Keeps the header row scrollable.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoFreeze { get; set; }

    /// <summary>Disable composer auto-formatting for dynamic collection columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoAutoFormatDynamicCollections { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    private readonly List<object?> _items = new();
    /// <inheritdoc />
    protected override void ProcessRecord() => TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    /// <inheritdoc />
    protected override void EndProcessing() {
        var context = ExcelDslContext.Require(this);
        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var data = ExcelTabularInputService.ToDataTable(rows, copyExistingTables: false, normalizerOptions: ExcelFlattenOptionsService.Create(FlattenOptions, CollectionSeparator, null, (name, error) => WriteWarning($"Skipping property '{name}' because it could not be read: {error.Message}")));
        var sourceColumns = data.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName).ToArray();
        var created = context.RequireComposer().AddTable(data, Title, configure: options => {
            options.Columns = NormalizePropertyList(Property) ?? sourceColumns;
            options.ExcludeProperties = NormalizePropertyList(ExcludeProperty) ?? Array.Empty<string>();
        }, style: TableStyle, autoFilter: !NoAutoFilter, freezeHeaderRow: !NoFreeze, visuals: options => {
            options.AutoFormatDynamicCollections = !NoAutoFormatDynamicCollections;
            options.ShowFirstColumn = ShowFirstColumn;
            options.ShowLastColumn = ShowLastColumn;
            options.ShowRowStripes = !NoRowStripes;
            options.ShowColumnStripes = ShowColumnStripes;
        });
        if (Name != null) created = created.Rename(Name);
        context.LastReportTable = created;
        var plan = ExcelColumnFormatPlanService.Build(ColumnFormat, null, null, null, null, null, null, null, 2, null);
        if (plan != null) created.Sheet.ApplyColumnFormatPlan(plan, headerRow: ExcelA1.ParseRange(created.Range!).Item1);
        if (PassThru) WriteObject(created);
    }
    private static string[]? NormalizePropertyList(string[]? properties) {
        if (properties == null || properties.Length == 0) {
            return null;
        }

        var normalized = properties
            .Where(static property => !string.IsNullOrWhiteSpace(property))
            .Select(static property => property.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return normalized.Length == 0 ? null : normalized;
    }
}
