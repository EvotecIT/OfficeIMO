using System;
using System.Collections;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes tabular data to the current worksheet and formats it as an Excel table.</summary>
/// <para>Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.</para>
/// <example>
/// <summary>Write objects as a table</summary>
/// <code>New-OfficeExcel -Path .\Services.xlsx {
///     ExcelSheet 'Services' {
///         ExcelTable (Get-Service)
///     }
/// }</code>
/// <para>Writes the services as a themed table with a frozen header, filters and fitted columns.</para>
/// </example>
/// <example>
/// <summary>Place two named tables on a sheet</summary>
/// <code>$targets = Import-OfficeCsv -Path .\Data\Q3-Targets.csv
/// $actuals = Import-OfficeCsv -Path .\Data\Q3-Actuals.csv
/// New-OfficeExcel -Path .\Reports\Q3-Performance.xlsx {
///     ExcelSheet 'Performance' {
///         ExcelTable $targets -Name Targets -Address A1
///         ExcelTable $actuals -Name Actuals -Address G1 -ColumnFormat @{ Attainment = 'Percent' }
///     }
/// }</code>
/// <para>Writes this quarter's targets and actuals side by side as named tables, so charts, formulas and later imports can refer to them by name.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTable", DefaultParameterSetName = "Default")]
[Alias("ExcelTable")]
public sealed class AddOfficeExcelTableCommand : PSCmdlet
{
    private readonly List<object?> _items = new();
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Name to assign to the table.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Name { get; set; }

    /// <summary>A1 address of the target cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Address { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Writes plain cells without creating a native Excel table.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoTable { get; set; }

    /// <summary>Creates the table without a named visual style.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoTableStyle { get; set; }

    /// <summary>Writes data without a header row.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Omits table filter buttons.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Retains native column widths without fitting them to the data.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoAutoFit { get; set; }

    /// <summary>Keeps the header row scrollable.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoFreeze { get; set; }

    /// <summary>Emphasizes the first table column; explicit false removes the emphasis.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasizes the last table column; explicit false removes the emphasis.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Enables alternating column bands; explicit false clears them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Disables alternating row bands. An explicit false value enables them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Maps column headers to number-format presets, format codes or per-column option hashtables.</summary>
    [Parameter(ParameterSetName = "Default")]
    public Hashtable? ColumnFormat { get; set; }

    /// <summary>Culture used for value conversion and number-format presets.</summary>
    [Parameter(ParameterSetName = "Default")]
    public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Text inserted between collection values flattened into one cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Nested-value projection policy; accepts a native options object or hashtable.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeOptionsTransformation(typeof(OfficeObjectFlattenOptions))]
    public OfficeObjectFlattenOptions? FlattenOptions { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        var dataSet = ExcelTabularInputService.TryGetSingleDataSet(_items);
        if (dataSet != null) {
            if (Sheet != null || Address != null || Name != null || View != OfficeTableView.Normal)
                throw new PSArgumentException("A DataSet creates one worksheet per DataTable. Sheet, Address, Name and View apply to single-table input.");
            var plan = ExcelColumnFormatPlanService.Build(ColumnFormat, null, null, null, null, null, null, null, 2, Culture.Name);
            if (plan != null) {
                if (NoHeader) throw new PSArgumentException("Column formats require a header row.");
                var headers = new HashSet<string>(dataSet.Tables.Cast<System.Data.DataTable>().SelectMany(table => table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName)), StringComparer.OrdinalIgnoreCase);
                foreach (var rule in plan.Rules) if (!headers.Contains(rule.Header)) throw new PSArgumentException($"Column '{rule.Header}' was not found in any DataSet table.");
            }
            var document = ExcelCommandTargets.Document(Document);
            var style = NoTableStyle ? ExcelTableStyle.None : TableStyle;
            foreach (var result in document.InsertDataSet(dataSet, createTables: !NoTable, tableStyle: style, includeHeaders: !NoHeader, includeAutoFilter: !NoAutoFilter, autoFit: !NoAutoFit)) {
                var target = document[result.SheetName];
                if (plan != null) target.ApplyColumnFormatPlan(plan, headerRow: 1);
                if (!NoTable && result.TableName != null)
                    ExcelTableStyleOptionService.Apply(target, result.TableName, style,
                        ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
                        ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
                        ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
                        ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
                if (!NoFreeze && !NoHeader) target.View.Freeze(1);
                if (PassThru) WriteObject(string.IsNullOrEmpty(result.Range) ? (object)result : NoTable ? (object)target.Range(result.Range) : result.TableName == null ? result : target.Table(result.TableName));
            }
            return;
        }
        var sheet = ResolveSheet();
        var (StartRow, StartColumn) = ExcelHostExtensions.ResolveCellAddress(null, null, Address ?? "A1");

        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var projectedRows = TableViewProjection.Project(rows, View);
        var normalizerOptions = ExcelFlattenOptionsService.Create(FlattenOptions, CollectionSeparator, Culture, (name, error) => WriteWarning($"Skipping property '{name}' because it could not be read: {error.Message}"));
        var table = ExcelTabularInputService.ToDataTable(
            projectedRows,
            Name,
            normalizerOptions: normalizerOptions);
        if (table.Columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer columns from the supplied data.");
        }

        if (StartRow < 1 || StartColumn < 1)
        {
            throw new ArgumentOutOfRangeException("StartRow/StartColumn must be 1 or greater.");
        }

        var formatPlan = ExcelColumnFormatPlanService.Build(ColumnFormat, null, null, null, null, null, null, null, 2, Culture.Name);
        if (formatPlan != null) {
            foreach (var rule in formatPlan.Rules) {
                if (!table.Columns.Contains(rule.Header))
                    throw new PSArgumentException($"Column '{rule.Header}' was not found in the supplied table.", nameof(ColumnFormat));
            }
        }
        var resolvedName = ResolveName(table);
        ExcelTable? created = null;
        string range;
        if (NoTable) { sheet.WriteData(table, StartRow, StartColumn, includeHeaders: !NoHeader); range = ExcelA1.CellReference(StartRow, StartColumn) + ":" + ExcelA1.CellReference(StartRow + Math.Max(1, table.Rows.Count + (NoHeader ? 0 : 1)) - 1, StartColumn + table.Columns.Count - 1); }
        else { created = sheet.AddTable(table, new ExcelTableOptions { AutoFit = false, UseDefaultHeaderlessColumnNames = true, IncludeHeaders = !NoHeader, Name = resolvedName, Style = NoTableStyle ? ExcelTableStyle.None : TableStyle, IncludeAutoFilter = !NoAutoFilter }, row: StartRow, column: StartColumn); range = created.Range!; }
        if (created != null) ExcelTableStyleOptionService.Apply(
            sheet,
            range,
            NoTableStyle ? ExcelTableStyle.None : TableStyle,
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
        var context = ExcelDslContext.Current;
        context?.RegisterTableRange(sheet, resolvedName, range);

        if (formatPlan != null && table.Rows.Count > 0) {
            int firstDataRow = StartRow + (NoHeader.IsPresent ? 0 : 1);
            foreach (var rule in formatPlan.Rules) {
                int column = StartColumn + table.Columns[rule.Header]!.Ordinal;
                int firstRow = rule.IncludeHeader ? StartRow : firstDataRow;
                sheet.Range(ExcelA1.CellReference(firstRow, column) + ":" +
                    ExcelA1.CellReference(firstDataRow + table.Rows.Count - 1, column)).SetNumberFormat(rule.ResolveNumberFormat());
                if (rule.AutoFit) sheet.AutoFitColumn(column);
            }
        }

        if (!NoAutoFit.IsPresent)
        {
            var columnIndexes = Enumerable.Range(StartColumn, table.Columns.Count);
            sheet.AutoFitColumnsFor(columnIndexes);
        }

        if (!NoFreeze && !NoHeader) sheet.View.Freeze(StartRow);

        if (PassThru.IsPresent)
        {
            WriteObject(created != null ? (object)created : sheet.Range(range));
        }
    }

    private string? ResolveName(System.Data.DataTable table)
    {
        if (!string.IsNullOrWhiteSpace(Name))
        {
            return Name;
        }

        return string.IsNullOrWhiteSpace(table.TableName) ? null : table.TableName;
    }

    private ExcelSheet ResolveSheet() => ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
}
