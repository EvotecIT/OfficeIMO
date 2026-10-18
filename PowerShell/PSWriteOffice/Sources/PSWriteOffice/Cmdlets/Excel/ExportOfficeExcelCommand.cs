using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Exports PowerShell objects to an Excel workbook using an operator-friendly surface.</summary>
/// <para>Provides a fast PowerShell export path while keeping OfficeIMO as the workbook engine.</para>
/// <example>
///   <summary>Export objects to a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Data -AutoFit -FreezeTopRow</code>
///   <para>Creates a workbook, writes the objects as a table, auto-fits columns, and freezes the header row.</para>
/// </example>
/// <example>
///   <summary>Export objects with report-friendly column formats.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Sales -TextColumn Id -CurrencyColumn Revenue -ColumnFormat @{ Rate = @{ Style = 'Percent'; Decimals = 1 }; Created = 'Date' } -FormatCultureName en-US -AutoFitFormattedColumn</code>
///   <para>Formats ID values as text, Revenue as currency, Rate as a one-decimal percentage, and Created as a short date while keeping formatting logic in OfficeIMO.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeExcel", DefaultParameterSetName = ParameterSetCreate, SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium)]
[Alias("ExcelExport")]
public sealed partial class ExportOfficeExcelCommand : PSCmdlet
{
    private const string ParameterSetCreate = "Create";
    private const string ParameterSetAppend = "Append";
    private const string ParameterSetClearSheet = "ClearSheet";
    private readonly List<object?> _input = new();

    /// <summary>Destination workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Objects to write. Accepts pipeline input.</summary>
    [Parameter(ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Worksheet name to create or update.</summary>
    [Parameter]
    [Alias("Sheet")]
    public string WorksheetName { get; set; } = "Sheet1";

    /// <summary>Optional Excel table name.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Built-in Excel table style name.</summary>
    [Parameter]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Emphasize the first table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasize the last table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Disable alternating row stripes for newly created tables.</summary>
    [Parameter]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Enable alternating column stripes for newly created tables.</summary>
    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Starting row for new exports. When appending and left at 1, rows are written after the used range.</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Starting column for new exports.</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Do not emit a header row.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Do not create an Excel table around the exported data.</summary>
    [Parameter]
    public SwitchParameter NoTable { get; set; }

    /// <summary>Disable AutoFilter dropdowns on the created table.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Auto-fit exported columns.</summary>
    [Parameter]
    [Alias("AutoSize")]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Freeze the exported header row.</summary>
    [Parameter]
    public SwitchParameter FreezeTopRow { get; set; }

    /// <summary>Freeze the first exported column.</summary>
    [Parameter]
    public SwitchParameter FreezeFirstColumn { get; set; }

    /// <summary>Bold the exported header row.</summary>
    [Parameter]
    public SwitchParameter BoldTopRow { get; set; }

    /// <summary>Write a title above the exported table.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Append rows to an existing worksheet when the workbook exists.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetAppend)]
    public SwitchParameter Append { get; set; }

    /// <summary>Require append operations to extend an existing Excel table instead of writing after the used range.</summary>
    [Parameter(ParameterSetName = ParameterSetAppend)]
    public SwitchParameter AppendToTable { get; set; }

    /// <summary>Replace the target worksheet inside an existing workbook.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetClearSheet)]
    public SwitchParameter ClearSheet { get; set; }

    /// <summary>Do not overwrite an existing workbook unless appending or clearing a sheet.</summary>
    [Parameter(ParameterSetName = ParameterSetCreate)]
    [Parameter(ParameterSetName = ParameterSetAppend)]
    [Parameter(ParameterSetName = ParameterSetClearSheet)]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Exclude specific properties from exported objects.</summary>
    [Parameter]
    public string[]? ExcludeProperty { get; set; }

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

    /// <summary>Header-to-format map. Values may be preset names such as Text, Currency, Percent, Date, or custom Excel number formats.</summary>
    [Parameter]
    public Hashtable? ColumnFormat { get; set; }

    /// <summary>Headers that should be formatted as text, useful for IDs, zip codes, and leading-zero values.</summary>
    [Parameter]
    public string[]? TextColumn { get; set; }

    /// <summary>Headers that should be formatted as decimal numbers.</summary>
    [Parameter]
    public string[]? NumberColumn { get; set; }

    /// <summary>Headers that should be formatted as whole numbers.</summary>
    [Parameter]
    public string[]? IntegerColumn { get; set; }

    /// <summary>Headers that should be formatted as percentages.</summary>
    [Parameter]
    public string[]? PercentColumn { get; set; }

    /// <summary>Headers that should be formatted as currency.</summary>
    [Parameter]
    public string[]? CurrencyColumn { get; set; }

    /// <summary>Headers that should be formatted as dates.</summary>
    [Parameter]
    public string[]? DateColumn { get; set; }

    /// <summary>Headers that should be formatted as date/time values.</summary>
    [Parameter]
    public string[]? DateTimeColumn { get; set; }

    /// <summary>Decimal places used by friendly number, percent, and currency column presets.</summary>
    [Parameter]
    public int FormatDecimals { get; set; } = 2;

    /// <summary>Culture used by friendly currency column presets, such as en-US or pl-PL.</summary>
    [Parameter]
    public string? FormatCultureName { get; set; }

    /// <summary>Include header cells when applying export-time column formats.</summary>
    [Parameter]
    public SwitchParameter IncludeHeaderInColumnFormat { get; set; }

    /// <summary>Auto-fit only columns that receive export-time column formats.</summary>
    [Parameter]
    public SwitchParameter AutoFitFormattedColumn { get; set; }

    /// <summary>Continue when a requested export-time column format header is missing.</summary>
    [Parameter]
    public SwitchParameter IgnoreMissingColumnFormat { get; set; }

    /// <summary>Include properties that cannot be read by exporting a descriptive placeholder value.</summary>
    [Parameter]
    public SwitchParameter IncludeUnexportableProperties { get; set; }

    /// <summary>Controls how unreadable PowerShell properties are handled while projecting export rows.</summary>
    [Parameter]
    public ActionPreference PropertyConversionErrorAction { get; set; } = ActionPreference.SilentlyContinue;

    /// <summary>Open the workbook after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Workbook document title metadata.</summary>
    [Parameter]
    public string? DocumentTitle { get; set; }

    /// <summary>Workbook author metadata.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Workbook subject metadata.</summary>
    [Parameter]
    public string? Subject { get; set; }

    /// <summary>Workbook keyword metadata.</summary>
    [Parameter]
    public string? Keywords { get; set; }

    /// <summary>Workbook description metadata.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Workbook category metadata.</summary>
    [Parameter]
    public string? Category { get; set; }

    /// <summary>Workbook company metadata.</summary>
    [Parameter]
    public string? Company { get; set; }

    /// <summary>Workbook manager metadata.</summary>
    [Parameter]
    public string? Manager { get; set; }

    /// <summary>Workbook application-name metadata.</summary>
    [Parameter]
    public string? ApplicationName { get; set; }

    /// <summary>Workbook last-modified-by metadata.</summary>
    [Parameter]
    public string? LastModifiedBy { get; set; }

    /// <summary>Run OfficeIMO worksheet preflight cleanup before saving.</summary>
    [Parameter]
    public SwitchParameter SafePreflight { get; set; }

    /// <summary>Repair common defined-name issues before saving.</summary>
    [Parameter]
    public SwitchParameter SafeRepairDefinedNames { get; set; }

    /// <summary>Validate the saved package with OpenXmlValidator and throw on errors.</summary>
    [Parameter]
    public SwitchParameter ValidateOpenXml { get; set; }

    /// <summary>Disable OfficeIMO fast package writers for this save.</summary>
    [Parameter]
    public SwitchParameter DisableFastPackageWriter { get; set; }

    /// <summary>Evaluate supported formulas and write cached values before saving.</summary>
    [Parameter]
    public SwitchParameter EvaluateFormulas { get; set; }

    /// <summary>Remove cached formula results before saving.</summary>
    [Parameter]
    public SwitchParameter ClearCachedFormulaResults { get; set; }

    /// <summary>Mark formula cells dirty before saving.</summary>
    [Parameter]
    public SwitchParameter MarkFormulasDirty { get; set; }

    /// <summary>Request a full workbook recalculation when opened in Excel-compatible applications.</summary>
    [Parameter]
    public SwitchParameter ForceFullCalculationOnOpen { get; set; }

    /// <summary>Workbook date system for Excel date serials.</summary>
    [Parameter]
    [ValidateSet("1900", "1904", "NineteenHundred", "NineteenFour")]
    public string? DateSystem { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        AddInput(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_input.Count == 0)
        {
            throw new PSArgumentException("Provide at least one object to export.", nameof(InputObject));
        }

        if (StartRow < 1 || StartColumn < 1)
        {
            throw new PSArgumentException("StartRow and StartColumn must be 1 or greater.");
        }

        if (Append.IsPresent && ClearSheet.IsPresent)
        {
            throw new PSArgumentException("Specify either -Append or -ClearSheet, but not both.");
        }

        if (AppendToTable.IsPresent && !Append.IsPresent)
        {
            throw new PSArgumentException("Use -AppendToTable together with -Append.");
        }

        if (AppendToTable.IsPresent && NoTable.IsPresent)
        {
            throw new PSArgumentException("Use either -AppendToTable or -NoTable, but not both.");
        }

        var style = TableStyle;

        var columnFormatPlan = BuildColumnFormatPlan();

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (File.Exists(resolvedPath) && NoClobber.IsPresent && !Append.IsPresent && !ClearSheet.IsPresent)
        {
            throw new IOException($"File '{resolvedPath}' already exists.");
        }

        var preserveWorkbook = File.Exists(resolvedPath) && (Append.IsPresent || ClearSheet.IsPresent);
        var action = preserveWorkbook
            ? Append.IsPresent ? "Update existing Excel workbook" : "Replace worksheet in existing Excel workbook"
            : File.Exists(resolvedPath) ? "Overwrite Excel workbook" : "Write new Excel workbook";
        if (!ShouldProcess(resolvedPath, action))
        {
            return;
        }

        var saveOptions = ExcelDocumentService.CreateSaveOptions(
            SafePreflight.IsPresent,
            SafeRepairDefinedNames.IsPresent,
            ValidateOpenXml.IsPresent,
            DisableFastPackageWriter.IsPresent,
            EvaluateFormulas.IsPresent,
            ClearCachedFormulaResults.IsPresent,
            MarkFormulasDirty.IsPresent,
            ForceFullCalculationOnOpen.IsPresent);

        var directory = System.IO.Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var directReader = ExcelTabularInputService.TryGetSingleDataReader(_input);
        if (directReader != null && TryWriteDataReaderPackageDirectly(
                directReader,
                resolvedPath,
                preserveWorkbook,
                style,
                columnFormatPlan))
        {
            return;
        }

        if (File.Exists(resolvedPath) && !preserveWorkbook)
        {
            File.Delete(resolvedPath);
        }

        var document = preserveWorkbook
            ? ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false)
            : ExcelDocumentService.CreateDocument(resolvedPath, autoSave: false);

        try
        {
            ExcelDateSystemService.ApplyIfSpecified(document, DateSystem, nameof(DateSystem));
            ExcelDocumentPropertyService.ApplyCommonProperties(
                document,
                DocumentTitle,
                Author,
                Subject,
                Keywords,
                Description,
                Category,
                Company,
                Manager,
                ApplicationName,
                LastModifiedBy);

            var dataSet = ExcelTabularInputService.TryGetSingleDataSet(_input);
            if (dataSet != null)
            {
                ExportDataSet(document, dataSet, style, columnFormatPlan);
                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, saveOptions: saveOptions);
                WritePassThru(resolvedPath);
                return;
            }

            var isAppendingToExistingSheet = Append.IsPresent && SheetExists(document, WorksheetName);
            if (AppendToTable.IsPresent && !isAppendingToExistingSheet)
            {
                throw new PSArgumentException($"Worksheet '{WorksheetName}' must exist when using -AppendToTable.");
            }

            var reader = ExcelTabularInputService.TryGetSingleDataReader(_input);
            if (reader != null && CanExportReaderDirectly(isAppendingToExistingSheet))
            {
                var readerSheet = PrepareSheet(document);
                var readerRange = ExportDataReader(
                    readerSheet,
                    new PowerShellNormalizingDataReader(reader, CreateNormalizerOptions()),
                    style,
                    columnFormatPlan);
                if (!string.IsNullOrWhiteSpace(readerRange))
                {
                    WriteVerbose($"Exported data reader to {readerSheet.Name}!{readerRange}.");
                }

                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, saveOptions: saveOptions);
                WritePassThru(resolvedPath);
                return;
            }

            var sheet = PrepareSheet(document);
            var dataStartRow = ResolveDataStartRow(sheet, isAppendingToExistingSheet);
            var includeHeaders = !NoHeader.IsPresent && !isAppendingToExistingSheet;

            if (TryExportObjectsThroughOfficeImoObjectPath(
                sheet,
                isAppendingToExistingSheet,
                preserveWorkbook,
                dataStartRow,
                includeHeaders,
                style,
                columnFormatPlan,
                out var objectRange))
            {
                if (!string.IsNullOrWhiteSpace(objectRange))
                {
                    WriteVerbose($"Exported data to {sheet.Name}!{objectRange}.");
                }

                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, saveOptions: saveOptions);
                WritePassThru(resolvedPath);
                return;
            }

            var data = BuildDataTable();
            if (data.Columns.Count == 0)
            {
                throw new InvalidOperationException("Unable to infer columns from the supplied data.");
            }

            if (!string.IsNullOrWhiteSpace(Title) && !isAppendingToExistingSheet)
            {
                sheet.Cell(dataStartRow, StartColumn, Title!);
                sheet.CellBold(dataStartRow, StartColumn, true);
                dataStartRow++;
            }

            var appendTableName = ResolveAppendTableName(document, sheet, isAppendingToExistingSheet);
            var range = WriteData(sheet, data, dataStartRow, includeHeaders, style, isAppendingToExistingSheet, appendTableName, ResolveTableName(data));
            ApplyColumnFormatPlan(
                sheet,
                columnFormatPlan,
                ResolveColumnFormatHeaderRow(document, sheet, columnFormatPlan, isAppendingToExistingSheet, appendTableName, includeHeaders, dataStartRow),
                requireExplicitHeaderRow: !includeHeaders);

            if (BoldTopRow.IsPresent && includeHeaders)
            {
                BoldRow(sheet, dataStartRow, StartColumn, data.Columns.Count);
            }

            if (AutoFit.IsPresent)
            {
                sheet.AutoFitColumnsFor(Enumerable.Range(StartColumn, data.Columns.Count));
            }

            if (FreezeTopRow.IsPresent || FreezeFirstColumn.IsPresent)
            {
                var frozenRows = ResolveFrozenTopRows(document, sheet, isAppendingToExistingSheet, appendTableName, dataStartRow, includeHeaders);
                var frozenColumns = FreezeFirstColumn.IsPresent ? Math.Max(1, StartColumn) : 0;
                sheet.Freeze(frozenRows, frozenColumns);
            }

            if (!string.IsNullOrWhiteSpace(range))
            {
                WriteVerbose($"Exported data to {sheet.Name}!{range}.");
            }

            ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, saveOptions: saveOptions);
        }
        catch
        {
            ExcelDocumentService.CloseDocument(document);
            throw;
        }

        if (PassThru.IsPresent)
        {
            WritePassThru(resolvedPath);
        }
    }

    private bool CanExportReaderDirectly(bool isAppendingToExistingSheet)
    {
        return !isAppendingToExistingSheet && ExcludeProperty is not { Length: > 0 };
    }

    private bool TryExportObjectsThroughOfficeImoObjectPath(
        ExcelSheet sheet,
        bool isAppendingToExistingSheet,
        bool preserveWorkbook,
        int dataStartRow,
        bool includeHeaders,
        ExcelTableStyle style,
        ExcelColumnFormatPlan? columnFormatPlan,
        out string range)
    {
        range = string.Empty;
        if (!CanExportObjectsThroughOfficeImoObjectPath(isAppendingToExistingSheet, preserveWorkbook, dataStartRow, includeHeaders, columnFormatPlan))
        {
            return false;
        }

        var items = _input.Where(static item => item != null).ToList();
        if (items.Count == 0)
        {
            return false;
        }

        if (!CanUseOfficeImoObjectPathInput(items))
        {
            return false;
        }

        var normalizedItems = PowerShellObjectNormalizer
            .NormalizeItems(items, CreateNormalizerOptions())
            .ToList();
        sheet.InsertObjects<object?>(normalizedItems, includeHeaders, dataStartRow);
        range = sheet.UsedRangeA1;
        if (string.IsNullOrWhiteSpace(range))
        {
            return false;
        }

        if (!NoTable.IsPresent)
        {
            sheet.AddTable(range, includeHeaders, TableName ?? string.Empty, style, includeAutoFilter: !NoAutoFilter.IsPresent);
            ApplyTableStyleOptions(sheet, range, style);
        }

        ApplyColumnFormatPlan(sheet, columnFormatPlan, includeHeaders ? dataStartRow : null);
        return true;
    }

    private bool CanExportObjectsThroughOfficeImoObjectPath(
        bool isAppendingToExistingSheet,
        bool preserveWorkbook,
        int dataStartRow,
        bool includeHeaders,
        ExcelColumnFormatPlan? columnFormatPlan)
    {
        return !isAppendingToExistingSheet &&
            !preserveWorkbook &&
            includeHeaders &&
            dataStartRow == 1 &&
            StartColumn == 1 &&
            string.IsNullOrWhiteSpace(Title) &&
            ExcludeProperty is not { Length: > 0 } &&
            !AutoFit.IsPresent &&
            !AutoFitFormattedColumn.IsPresent &&
            !ColumnFormatRequiresMaterializedCells(columnFormatPlan) &&
            !BoldTopRow.IsPresent &&
            !FreezeTopRow.IsPresent &&
            !FreezeFirstColumn.IsPresent;
    }

    private static bool CanUseOfficeImoObjectPathInput(IReadOnlyList<object?> items)
    {
        foreach (var item in items)
        {
            if (item is PSObject)
            {
                return false;
            }

            if (item is Dictionary<string, object?> ||
                item is IReadOnlyDictionary<string, object?> ||
                item is IDictionary<string, object?> ||
                item is IDictionary)
            {
                continue;
            }

            return false;
        }

        return items.Count > 0;
    }

    private ExcelColumnFormatPlan? BuildColumnFormatPlan()
        => ExcelColumnFormatPlanService.Build(
            ColumnFormat,
            TextColumn,
            NumberColumn,
            IntegerColumn,
            PercentColumn,
            CurrencyColumn,
            DateColumn,
            DateTimeColumn,
            FormatDecimals,
            FormatCultureName);

    private IReadOnlyList<ExcelColumnFormatResult> ApplyColumnFormatPlan(
        ExcelSheet sheet,
        ExcelColumnFormatPlan? plan,
        int? headerRow = null,
        bool throwOnMissing = true,
        bool requireExplicitHeaderRow = false)
    {
        if (plan == null)
        {
            return Array.Empty<ExcelColumnFormatResult>();
        }

        if (requireExplicitHeaderRow && !headerRow.HasValue)
        {
            var headers = string.Join(", ", GetColumnFormatHeaders(plan));
            WriteVerbose($"Column format headers were not applied on worksheet '{sheet.Name}' because no header row was emitted.");
            if (throwOnMissing && !IgnoreMissingColumnFormat.IsPresent)
            {
                throw new PSArgumentException($"Column format headers were not found on worksheet '{sheet.Name}': {headers}. Export-time column formats require a header row; remove -NoHeader or use -IgnoreMissingColumnFormat for optional columns.");
            }

            return Array.Empty<ExcelColumnFormatResult>();
        }

        var results = sheet.ApplyColumnFormatPlan(
            plan,
            includeHeader: IncludeHeaderInColumnFormat.IsPresent,
            autoFit: AutoFitFormattedColumn.IsPresent,
            headerRow: headerRow);
        var missing = results.Where(static result => !result.Applied).ToArray();
        foreach (var result in missing)
        {
            WriteVerbose(result.Warning);
        }

        if (missing.Length > 0 && throwOnMissing && !IgnoreMissingColumnFormat.IsPresent)
        {
            var headers = string.Join(", ", missing.Select(static result => result.Header));
            throw new PSArgumentException($"Column format headers were not found on worksheet '{sheet.Name}': {headers}. Use -IgnoreMissingColumnFormat for optional columns.");
        }

        return results;
    }

    private static string[] GetColumnFormatHeaders(ExcelColumnFormatPlan plan)
        => plan.Rules
            .Select(static rule => rule.Header.Trim())
            .Where(static header => header.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private bool ColumnFormatRequiresMaterializedCells(ExcelColumnFormatPlan? plan)
    {
        return IncludeHeaderInColumnFormat.IsPresent || (plan?.Rules.Any(static rule => rule.IncludeHeader) ?? false);
    }

    private string ExportDataReader(ExcelSheet sheet, IDataReader reader, ExcelTableStyle style, ExcelColumnFormatPlan? columnFormatPlan)
    {
        var dataStartRow = StartRow;
        var fieldCount = reader.FieldCount;
        if (!string.IsNullOrWhiteSpace(Title))
        {
            sheet.Cell(dataStartRow, StartColumn, Title!);
            sheet.CellBold(dataStartRow, StartColumn, true);
            dataStartRow++;
        }

        var range = sheet.InsertDataReader(
            reader,
            startRow: dataStartRow,
            startColumn: StartColumn,
            includeHeaders: !NoHeader.IsPresent,
            tableName: TableName,
            style: style,
            includeAutoFilter: !NoAutoFilter.IsPresent,
            createTable: !NoTable.IsPresent,
            autoFit: false);

        if (!string.IsNullOrWhiteSpace(range))
        {
            ApplyColumnFormatPlan(sheet, columnFormatPlan, !NoHeader.IsPresent ? dataStartRow : null, requireExplicitHeaderRow: NoHeader.IsPresent);

            if (!NoTable.IsPresent)
            {
                ApplyTableStyleOptions(sheet, range, style);
            }

            if (BoldTopRow.IsPresent && !NoHeader.IsPresent)
            {
                BoldRow(sheet, dataStartRow, StartColumn, fieldCount);
            }

            if (AutoFit.IsPresent)
            {
                sheet.AutoFitColumnsFor(Enumerable.Range(StartColumn, fieldCount));
            }

            if (FreezeTopRow.IsPresent || FreezeFirstColumn.IsPresent)
            {
                var frozenRows = FreezeTopRow.IsPresent ? Math.Max(1, !NoHeader.IsPresent ? dataStartRow : StartRow) : 0;
                var frozenColumns = FreezeFirstColumn.IsPresent ? Math.Max(1, StartColumn) : 0;
                sheet.Freeze(frozenRows, frozenColumns);
            }
        }

        return range;
    }

    private void ExportDataSet(ExcelDocument document, DataSet dataSet, ExcelTableStyle style, ExcelColumnFormatPlan? columnFormatPlan)
    {
        if (dataSet.Tables.Count == 0)
        {
            throw new PSArgumentException("DataSet must contain at least one DataTable.", nameof(InputObject));
        }

        var sheetNames = BuildDataSetWorksheetNames(document, dataSet, Append.IsPresent || ClearSheet.IsPresent);
        var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var formattedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tableIndex = 1;
        foreach (DataTable sourceTable in dataSet.Tables)
        {
            var sheetName = sheetNames[tableIndex - 1];

            var sheetExists = SheetExists(document, sheetName);
            if (ClearSheet.IsPresent && sheetExists)
            {
                document.RemoveWorksheet(sheetName);
                sheetExists = false;
            }

            var sheet = GetOrAddDataSetSheet(document, sheetName);
            var table = PrepareDataTableForExport(sourceTable);
            if (table.Columns.Count == 0)
            {
                throw new InvalidOperationException($"Unable to infer columns from DataTable '{sourceTable.TableName}'.");
            }

            var isAppendingToExistingSheet = Append.IsPresent && sheetExists;
            var dataStartRow = ResolveDataStartRow(sheet, isAppendingToExistingSheet);
            var includeHeaders = !NoHeader.IsPresent && !isAppendingToExistingSheet;
            var appendTableName = ResolveAppendTableName(document, sheet, isAppendingToExistingSheet);
            var tableName = ResolveDataSetTableName(table, usedTableNames);
            var range = WriteData(sheet, table, dataStartRow, includeHeaders, style, isAppendingToExistingSheet, appendTableName, tableName);
            foreach (var result in ApplyColumnFormatPlan(
                         sheet,
                         columnFormatPlan,
                         ResolveColumnFormatHeaderRow(document, sheet, columnFormatPlan, isAppendingToExistingSheet, appendTableName, includeHeaders, dataStartRow),
                         throwOnMissing: false,
                         requireExplicitHeaderRow: !includeHeaders))
            {
                if (result.Applied)
                {
                    formattedHeaders.Add(result.Header);
                }
            }

            if (BoldTopRow.IsPresent && includeHeaders)
            {
                BoldRow(sheet, dataStartRow, StartColumn, table.Columns.Count);
            }

            if (AutoFit.IsPresent)
            {
                sheet.AutoFitColumnsFor(Enumerable.Range(StartColumn, table.Columns.Count));
            }

            if (FreezeTopRow.IsPresent || FreezeFirstColumn.IsPresent)
            {
                var frozenRows = ResolveFrozenTopRows(document, sheet, isAppendingToExistingSheet, appendTableName, dataStartRow, includeHeaders);
                var frozenColumns = FreezeFirstColumn.IsPresent ? Math.Max(1, StartColumn) : 0;
                sheet.Freeze(frozenRows, frozenColumns);
            }

            if (!string.IsNullOrWhiteSpace(range))
            {
                WriteVerbose($"Exported DataTable '{sourceTable.TableName}' to {sheet.Name}!{range}.");
            }

            tableIndex++;
        }

        ThrowForDataSetColumnFormatsMissingEverywhere(columnFormatPlan, formattedHeaders);
    }

    private void ThrowForDataSetColumnFormatsMissingEverywhere(ExcelColumnFormatPlan? plan, HashSet<string> formattedHeaders)
    {
        if (plan == null || IgnoreMissingColumnFormat.IsPresent)
        {
            return;
        }

        var missingEverywhere = plan.Rules
            .Select(static rule => rule.Header)
            .Where(header => !formattedHeaders.Contains(header))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (missingEverywhere.Length == 0)
        {
            return;
        }

        throw new PSArgumentException($"Column format headers were not found in any DataSet worksheet: {string.Join(", ", missingEverywhere)}. Use -IgnoreMissingColumnFormat for optional columns.");
    }

    private DataTable PrepareDataTableForExport(DataTable sourceTable)
    {
        var table = ExcelTabularInputService.ToDataTable(
            new object?[] { sourceTable },
            sourceTable.TableName,
            copyExistingTables: ExcludeProperty is { Length: > 0 },
            normalizerOptions: CreateNormalizerOptions());
        return ApplyExcludedColumns(table);
    }

    private string? ResolveDataSetTableName(DataTable table, ISet<string> usedTableNames)
    {
        var tableName = ResolveTableName(table, allowExplicitOverride: false);
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        var candidate = tableName!;
        var suffix = 2;
        while (!usedTableNames.Add(candidate))
        {
            candidate = $"{tableName}_{suffix}";
            suffix++;
        }

        return candidate;
    }

    private ExcelSheet PrepareSheet(ExcelDocument document)
    {
        if (ClearSheet.IsPresent && SheetExists(document, WorksheetName))
        {
            document.RemoveWorksheet(WorksheetName);
        }

        return document.GetOrCreateSheet(WorksheetName, ExcelSheetNameValidationMode.Sanitize);
    }

    private DataTable BuildDataTable()
    {
        var table = ExcelTabularInputService.ToDataTable(
            _input,
            TableName,
            copyExistingTables: ExcludeProperty is { Length: > 0 },
            normalizerOptions: CreateNormalizerOptions());
        return ApplyExcludedColumns(table);
    }

    private PowerShellObjectNormalizerOptions CreateNormalizerOptions()
    {
        return new PowerShellObjectNormalizerOptions
        {
            CollectionSeparator = CollectionSeparator,
            DictionaryEntrySeparator = DictionaryEntrySeparator,
            DictionaryKeyValueSeparator = DictionaryKeyValueSeparator,
            MaxCollectionItems = MaxCollectionItems,
            MaxNestingDepth = MaxNestingDepth,
            IncludeUnexportableProperties = IncludeUnexportableProperties.IsPresent,
            PropertyErrorAction = PropertyConversionErrorAction,
            PropertyErrorCallback = PropertyConversionErrorAction == ActionPreference.Continue
                ? (name, exception) => WriteWarning($"Skipping property '{name}' because it could not be read: {exception.Message}")
                : null,
            UnexportablePropertyValueFactory = static (_, exception) => $"Property export failed: {exception.Message}"
        };
    }

    private DataTable ApplyExcludedColumns(DataTable table)
    {
        if (ExcludeProperty is { Length: > 0 })
        {
            var excluded = new HashSet<string>(
                ExcludeProperty.Where(static p => !string.IsNullOrWhiteSpace(p)).Select(static p => p.Trim()),
                StringComparer.OrdinalIgnoreCase);

            foreach (DataColumn column in table.Columns.Cast<DataColumn>().ToArray())
            {
                if (excluded.Contains(column.ColumnName))
                {
                    table.Columns.Remove(column);
                }
            }
        }

        return table;
    }

    private void WritePassThru(string resolvedPath)
    {
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(resolvedPath));
        }
    }

    private string WriteData(ExcelSheet sheet, DataTable table, int startRow, bool includeHeaders, ExcelTableStyle style, bool appendRawRows, string? appendTableName, string? tableName)
    {
        if (appendRawRows &&
            !NoTable.IsPresent &&
            !string.IsNullOrWhiteSpace(appendTableName) &&
            TryAppendDataTableToTable(sheet, table, appendTableName!, out var updatedTableRange))
        {
            return updatedTableRange;
        }

        if (!NoTable.IsPresent && !appendRawRows)
        {
            var range = sheet.InsertDataTableAsTable(
                table,
                startRow,
                StartColumn,
                includeHeaders,
                tableName,
                style,
                includeAutoFilter: !NoAutoFilter.IsPresent);
            ApplyTableStyleOptions(sheet, range, style);
            return range;
        }

        var cellCount = checked((table.Rows.Count + (includeHeaders ? 1 : 0)) * table.Columns.Count);
        var cells = new List<(int Row, int Column, object Value)>(cellCount);
        var row = startRow;

        if (includeHeaders)
        {
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                cells.Add((row, StartColumn + columnIndex, table.Columns[columnIndex].ColumnName));
            }
            row++;
        }

        foreach (DataRow dataRow in table.Rows)
        {
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                var value = dataRow[columnIndex];
                cells.Add((row, StartColumn + columnIndex, value is DBNull ? string.Empty : value));
            }
            row++;
        }

        sheet.CellValues(cells);
        var endRow = Math.Max(startRow, row - 1);
        var endColumn = StartColumn + table.Columns.Count - 1;
        return $"{A1.CellReference(startRow, StartColumn)}:{A1.CellReference(endRow, endColumn)}";
    }

    private void ApplyTableStyleOptions(ExcelSheet sheet, string? tableOrRange, ExcelTableStyle style)
    {
        ExcelTableStyleOptionService.Apply(
            sheet,
            tableOrRange,
            style,
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
    }

    private string? ResolveAppendTableName(ExcelDocument document, ExcelSheet sheet, bool appendToExistingSheet)
    {
        if (!appendToExistingSheet)
        {
            if (AppendToTable.IsPresent)
            {
                throw new PSArgumentException($"Worksheet '{sheet.Name}' must exist when using -AppendToTable.");
            }

            return null;
        }

        if (NoTable.IsPresent)
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(TableName))
        {
            return TableName;
        }

        var sheetTables = document.GetTables()
            .Where(table => string.Equals(table.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (sheetTables.Count == 1)
        {
            return sheetTables[0].Name;
        }

        if (AppendToTable.IsPresent)
        {
            throw new PSArgumentException(sheetTables.Count == 0
                ? $"Worksheet '{sheet.Name}' does not contain a table to append to. Provide -TableName or omit -AppendToTable."
                : $"Worksheet '{sheet.Name}' contains multiple tables. Provide -TableName when using -AppendToTable.");
        }

        return null;
    }

    private int? ResolveColumnFormatHeaderRow(
        ExcelDocument document,
        ExcelSheet sheet,
        ExcelColumnFormatPlan? plan,
        bool appendToExistingSheet,
        string? appendTableName,
        bool includeHeaders,
        int dataStartRow)
    {
        if (includeHeaders)
        {
            return dataStartRow;
        }

        if (!appendToExistingSheet || plan == null)
        {
            return null;
        }

        var tableHeaderRow = ResolveExistingTableHeaderRow(document, sheet, appendTableName);
        if (NoHeader.IsPresent)
        {
            return tableHeaderRow;
        }

        return tableHeaderRow ?? FindExistingHeaderRow(sheet, plan, allowPartialMatch: IgnoreMissingColumnFormat.IsPresent);
    }

    private int? ResolveExistingTableHeaderRow(ExcelDocument document, ExcelSheet sheet, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        var table = document.GetTables().FirstOrDefault(candidate =>
            string.Equals(candidate.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase));
        if (table != null && table.HasHeaderRow && !string.IsNullOrWhiteSpace(table.Range) &&
            A1.TryParseRange(table.Range, out var tableStartRow, out _, out _, out _))
        {
            return Math.Max(1, tableStartRow);
        }

        return null;
    }

    private static int? FindExistingHeaderRow(ExcelSheet sheet, ExcelColumnFormatPlan plan, bool allowPartialMatch)
    {
        var requestedHeaders = GetColumnFormatHeaders(plan);
        if (requestedHeaders.Length == 0)
        {
            return null;
        }

        var usedRange = sheet.UsedRangeA1;
        if (!A1.TryParseRange(usedRange, out var firstRow, out var firstColumn, out var lastRow, out var lastColumn))
        {
            var (row, column) = A1.ParseCellRef(usedRange);
            firstRow = lastRow = row;
            firstColumn = lastColumn = column;
        }

        var headerRow = Math.Max(1, firstRow);
        var rowHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var column = Math.Max(1, firstColumn); column <= lastColumn; column++)
        {
            if (sheet.TryGetCellText(headerRow, column, out var text) && !string.IsNullOrWhiteSpace(text))
            {
                rowHeaders.Add(text.Trim());
            }
        }

        var hasRequestedHeader = allowPartialMatch
            ? requestedHeaders.Any(rowHeaders.Contains)
            : requestedHeaders.All(rowHeaders.Contains);

        return hasRequestedHeader ? headerRow : null;
    }

    private string? ResolveTableName(DataTable table, bool allowExplicitOverride = true)
    {
        if (allowExplicitOverride && !string.IsNullOrWhiteSpace(TableName))
        {
            return TableName;
        }

        return string.IsNullOrWhiteSpace(table.TableName) ? null : table.TableName;
    }

    private int ResolveFrozenTopRows(ExcelDocument document, ExcelSheet sheet, bool appendToExistingSheet, string? appendTableName, int dataStartRow, bool includeHeaders)
    {
        if (!FreezeTopRow.IsPresent)
        {
            return 0;
        }

        if (!appendToExistingSheet)
        {
            return Math.Max(1, includeHeaders ? dataStartRow : StartRow);
        }

        if (!string.IsNullOrWhiteSpace(appendTableName))
        {
            var table = document.GetTables().FirstOrDefault(candidate =>
                string.Equals(candidate.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(candidate.Name, appendTableName, StringComparison.OrdinalIgnoreCase));
            if (table != null && !string.IsNullOrWhiteSpace(table.Range) &&
                A1.TryParseRange(table.Range, out var tableStartRow, out _, out _, out _))
            {
                return Math.Max(1, tableStartRow);
            }
        }

        return Math.Max(1, StartRow);
    }

    private bool TryAppendDataTableToTable(ExcelSheet sheet, DataTable table, string tableName, out string range)
    {
        range = sheet.AppendDataTableToTable(table, tableName, matchColumnsByHeader: true);
        return true;
    }

    private int ResolveDataStartRow(ExcelSheet sheet, bool appendToExistingSheet)
    {
        if (!appendToExistingSheet || StartRow > 1)
        {
            return StartRow;
        }

        var usedRange = sheet.UsedRangeA1;
        if (A1.TryParseRange(usedRange, out _, out _, out var lastRow, out _))
        {
            return lastRow + 1;
        }

        var (row, _) = A1.ParseCellRef(usedRange);
        return row > 0 ? row + 1 : StartRow;
    }

    private void AddInput(object? value)
    {
        if (value == null)
        {
            return;
        }

        if (value is PSObject psObject)
        {
            value = psObject.BaseObject is PSCustomObject ? psObject : psObject.BaseObject;
        }

        if (value is DataSet || value is DataTable || value is DataView || value is IDataReader)
        {
            _input.Add(value);
            return;
        }

        if (value is IEnumerable enumerable &&
            value is not string &&
            !PowerShellDictionaryAdapter.IsDictionaryLike(value))
        {
            foreach (var item in enumerable)
            {
                if (item != null)
                {
                    _input.Add(item);
                }
            }
            return;
        }

        _input.Add(value);
    }

    private static void BoldRow(ExcelSheet sheet, int row, int startColumn, int columnCount)
    {
        for (var column = startColumn; column < startColumn + columnCount; column++)
        {
            sheet.CellBold(row, column, true);
        }
    }

    private static bool SheetExists(ExcelDocument document, string worksheetName)
    {
        var normalized = NormalizeWorksheetName(worksheetName);
        return document.Sheets.Any(sheet =>
            string.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sheet.Name, normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static string[] BuildDataSetWorksheetNames(ExcelDocument document, DataSet dataSet, bool allowExistingMatches)
    {
        var names = new List<string>(dataSet.Tables.Count);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var existing = new HashSet<string>(
            document.Sheets.Select(sheet => sheet.Name).Where(name => !string.IsNullOrWhiteSpace(name)),
            StringComparer.OrdinalIgnoreCase);
        var tableIndex = 1;
        foreach (DataTable table in dataSet.Tables)
        {
            var requested = string.IsNullOrWhiteSpace(table.TableName)
                ? $"Table{tableIndex}"
                : table.TableName;

            var normalized = NormalizeWorksheetNameInfo(requested);
            var candidate = allowExistingMatches
                ? FindExistingWorksheetName(document, requested, normalized.Name, used, allowDirectNormalizedMatch: !normalized.UsedDefaultFallback)
                : null;

            string candidateName;
            if (!string.IsNullOrWhiteSpace(candidate) && !used.Contains(candidate!))
            {
                candidateName = candidate!;
            }
            else
            {
                candidateName = normalized.Name;
                var suffix = 2;
                while (used.Contains(candidateName) || existing.Contains(candidateName))
                {
                    candidateName = AppendWorksheetSuffix(normalized.Name, suffix);
                    suffix++;
                }
            }

            used.Add(candidateName);
            names.Add(candidateName);
            tableIndex++;
        }

        return names.ToArray();
    }

    private static string NormalizeWorksheetName(string worksheetName)
    {
        return NormalizeWorksheetNameInfo(worksheetName).Name;
    }

    private static (string Name, bool UsedDefaultFallback) NormalizeWorksheetNameInfo(string worksheetName)
    {
        var baseName = (worksheetName ?? string.Empty).Trim().Trim('\'', ' ');
        var chars = new char[baseName.Length];
        for (var i = 0; i < baseName.Length; i++)
        {
            var ch = baseName[i];
            chars[i] = IsInvalidWorksheetNameCharacter(ch) ? '_' : ch;
        }

        var cleaned = new string(chars).Trim();
        var usedDefaultFallback = false;
        if (string.IsNullOrEmpty(cleaned))
        {
            cleaned = "Sheet1";
            usedDefaultFallback = true;
        }

        var name = cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
        return (name, usedDefaultFallback);
    }

    private static ExcelSheet GetOrAddDataSetSheet(ExcelDocument document, string sheetName)
    {
        var sheet = document.Sheets.FirstOrDefault(existing =>
            string.Equals(existing.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        return sheet ?? document.AddWorksheet(sheetName, ExcelSheetNameValidationMode.None);
    }

    private static bool IsInvalidWorksheetNameCharacter(char ch)
    {
        return ch is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ||
            char.IsControl(ch) ||
            char.IsSurrogate(ch);
    }

    private static string? FindExistingWorksheetName(
        ExcelDocument document,
        string requestedName,
        string normalizedName,
        ISet<string> usedNames,
        bool allowDirectNormalizedMatch)
    {
        var existing = document.Sheets
            .Select(sheet => sheet.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name) && !usedNames.Contains(name))
            .ToArray();

        var directMatch = existing.FirstOrDefault(name =>
            string.Equals(name, requestedName, StringComparison.OrdinalIgnoreCase) ||
            (allowDirectNormalizedMatch && string.Equals(name, normalizedName, StringComparison.OrdinalIgnoreCase)));
        if (!string.IsNullOrWhiteSpace(directMatch))
        {
            return directMatch;
        }

        return existing
            .Select(name => (Name: name, Suffix: GetWorksheetNameSuffix(name, normalizedName)))
            .Where(match => match.Suffix.HasValue)
            .OrderBy(match => match.Suffix!.Value)
            .Select(match => match.Name)
            .FirstOrDefault();
    }

    private static string AppendWorksheetSuffix(string worksheetName, int suffix)
    {
        var suffixText = $" ({suffix})";
        var maxBase = 31 - suffixText.Length;
        var basePart = worksheetName.Length > maxBase ? worksheetName.Substring(0, maxBase) : worksheetName;
        return basePart + suffixText;
    }

    private static bool IsSuffixedWorksheetName(string worksheetName, string normalizedName)
    {
        return GetWorksheetNameSuffix(worksheetName, normalizedName).HasValue;
    }

    private static int? GetWorksheetNameSuffix(string worksheetName, string normalizedName)
    {
        var close = worksheetName.LastIndexOf(')');
        var open = worksheetName.LastIndexOf(" (", StringComparison.Ordinal);
        if (close != worksheetName.Length - 1 || open < 0 || open >= close)
        {
            return null;
        }

        var suffixText = worksheetName.Substring(open + 2, close - open - 2);
        if (!int.TryParse(suffixText, out var suffix) || suffix < 2)
        {
            return null;
        }

        return string.Equals(worksheetName, AppendWorksheetSuffix(normalizedName, suffix), StringComparison.OrdinalIgnoreCase)
            ? suffix
            : null;
    }
}
