using System.Collections;
using System.Data;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO;
using System.Globalization;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Exports PowerShell objects to an Excel workbook using an operator-friendly surface.</summary>
/// <para>Provides a fast PowerShell export path while keeping OfficeIMO as the workbook engine.</para>
/// <example>
/// <summary>Export objects in one line</summary>
/// <code>Get-Process | Export-OfficeExcel -Path .\Processes.xlsx -Show</code>
/// <para>Writes the running processes as a styled table with a frozen header and fitted columns, and opens the file.</para>
/// </example>
/// <example>
/// <summary>Export a titled sales sheet</summary>
/// <code>$export = @{
///     Path         = '.\Reports\Q3-Sales.xlsx'
///     Sheet        = 'Orders'
///     TableName    = 'Orders'
///     Title        = 'Q3 orders'
///     ColumnFormat = @{ Margin = 'Percent' }
/// }
/// Import-OfficeCsv -Path .\Data\Q3-Orders.csv |
///     Select-Object OrderId, Customer, Region, OrderDate, Amount, Margin |
///     Export-OfficeExcel @export</code>
/// <para>Exports the Q3 orders as a named table under a title; dates and amounts are formatted from the data, and only the margin column is set explicitly as a percentage.</para>
/// </example>
/// <example>
/// <summary>Append a daily snapshot to a history table</summary>
/// <code>Get-Volume |
///     Where-Object DriveLetter |
///     Select-Object @{ Name = 'Date'; Expression = { Get-Date } }, DriveLetter, Size, SizeRemaining |
///     Export-OfficeExcel -Path .\Reports\Disk-Usage.xlsx -Sheet History -TableName DiskHistory -Append</code>
/// <para>Adds today's free-space readings to the DiskHistory table, so the workbook grows into a capacity history one run at a time.</para>
/// </example>
/// <example>
/// <summary>Refresh one sheet of a workbook</summary>
/// <code>Import-OfficeCsv -Path .\Exports\open-tickets.csv |
///     Export-OfficeExcel -Path .\Reports\Helpdesk.xlsx -Sheet 'Open tickets' -ReplaceSheet</code>
/// <para>Replaces only the Open tickets sheet with this morning's export and leaves the other sheets, charts and notes in the workbook untouched.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeExcel", SupportsShouldProcess = true, DefaultParameterSetName = "Create")]
public sealed partial class ExportOfficeExcelCommand : PSCmdlet, IDisposable
{
    private readonly List<object?> _input = new();

    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = null!;

    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Optional Excel table name.</summary>
    [Parameter()]
    public string? TableName { get; set; }

    /// <summary>Write a title above the exported table.</summary>
    [Parameter()]
    public string? Title { get; set; }

    /// <summary>A1 address of the target cell.</summary>
    [Parameter()]
    public string? Address { get; set; }

    /// <summary>Built-in Office theme, native theme object or a document whose visual theme is read.</summary>
    [Parameter()]
    [OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Built-in Excel table style name.</summary>
    [Parameter()]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Writes plain cells without creating a native Excel table.</summary>
    [Parameter()]
    public SwitchParameter NoTable { get; set; }

    /// <summary>Creates the table without a named visual style.</summary>
    [Parameter()]
    public SwitchParameter NoTableStyle { get; set; }

    /// <summary>Writes data without a header row.</summary>
    [Parameter()]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Omits table filter buttons.</summary>
    [Parameter()]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Retains native column widths without fitting them to the data.</summary>
    [Parameter()]
    public SwitchParameter NoAutoFit { get; set; }

    /// <summary>Keeps the header row scrollable.</summary>
    [Parameter()]
    public SwitchParameter NoFreeze { get; set; }

    /// <summary>Freeze the first exported column.</summary>
    [Parameter()]
    public SwitchParameter FreezeFirstColumn { get; set; }

    /// <summary>Emphasizes the first table column; explicit false removes the emphasis.</summary>
    [Parameter()]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasizes the last table column; explicit false removes the emphasis.</summary>
    [Parameter()]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Enables alternating column bands; explicit false clears them.</summary>
    [Parameter()]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Disables alternating row bands. An explicit false value enables them.</summary>
    [Parameter()]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Maps column headers to number-format presets, format codes or per-column option hashtables.</summary>
    [Parameter()]
    public Hashtable? ColumnFormat { get; set; }

    /// <summary>Decimal places used for decimal, percent, currency, and scientific presets.</summary>
    [Parameter()]
    [ValidateRange(0, 30)]
    public int Decimals { get; set; } = 2;

    /// <summary>Culture used for value conversion and number-format presets.</summary>
    [Parameter()]
    public CultureInfo? Culture { get; set; }

    /// <summary>Continue when a requested export-time column format header is missing.</summary>
    [Parameter()]
    public SwitchParameter IgnoreMissingColumnFormat { get; set; }

    /// <summary>Exclude specific properties from exported objects.</summary>
    [Parameter()]
    public string[]? ExcludeProperty { get; set; }

    /// <summary>Text inserted between collection values flattened into one cell.</summary>
    [Parameter()]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Nested-value projection policy; accepts a native options object or hashtable.</summary>
    [Parameter()]
    [OfficeOptionsTransformation(typeof(OfficeObjectFlattenOptions))]
    public OfficeObjectFlattenOptions? FlattenOptions { get; set; }

    /// <summary>Workbook metadata keyed by native property name, such as Title, Author or Company.</summary>
    [Parameter()]
    public Hashtable? DocumentProperty { get; set; }

    /// <summary>Workbook date system for Excel date serials.</summary>
    [Parameter()]
    public ExcelDateSystem? DateSystem { get; set; }

    /// <summary>Native save options or a hashtable, applied to a completed content block.</summary>
    [Parameter()]
    [OfficeExcelSaveOptionsTransformation]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Open the saved file in its registered application after a local block completes.</summary>
    [Parameter()]
    public SwitchParameter Show { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Do not overwrite an existing workbook.</summary>
    [Parameter(ParameterSetName = "Create")]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Append rows to an existing worksheet when the workbook exists.</summary>
    [Parameter(ParameterSetName = "Append", Mandatory = true)]
    public SwitchParameter Append { get; set; }

    /// <summary>Replaces the selected worksheet while retaining other workbook content.</summary>
    [Parameter(ParameterSetName = "ReplaceSheet", Mandatory = true)]
    public SwitchParameter ReplaceSheet { get; set; }

    private string? _resolvedSheetName;
    private string WorksheetName => _resolvedSheetName ?? Sheet?.Name ?? (Sheet?.Sheet as ExcelSheet)?.Name ?? "Sheet1";
    private int StartRow => ExcelHostExtensions.ResolveCellAddress(null, null, Address ?? "A1").Row;
    private int StartColumn => ExcelHostExtensions.ResolveCellAddress(null, null, Address ?? "A1").Column;
    private SwitchParameter AutoFit => new SwitchParameter(!NoAutoFit);
    private SwitchParameter FreezeTopRow => new SwitchParameter(!NoFreeze);
    private SwitchParameter BoldTopRow => new SwitchParameter(!NoHeader);
    private bool AppendsToNamedTable => Append && !string.IsNullOrWhiteSpace(TableName);
    private SwitchParameter Open => Show;
    private SwitchParameter AutoFitFormattedColumn => new SwitchParameter(!NoAutoFit);
    private string DictionaryEntrySeparator => FlattenOptions?.DictionaryEntrySeparator ?? "; ";
    private string DictionaryKeyValueSeparator => FlattenOptions?.DictionaryKeyValueSeparator ?? ": ";
    private int MaxCollectionItems => FlattenOptions?.MaxCollectionItems ?? PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;
    private int MaxNestingDepth => FlattenOptions?.MaxNestingDepth ?? PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;
    private SwitchParameter IncludeUnexportableProperties => new SwitchParameter(FlattenOptions?.IncludeUnreadableProperties ?? false);
    private ActionPreference PropertyConversionErrorAction => FlattenOptions?.PropertyConversionErrorAction ?? ActionPreference.SilentlyContinue;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!StreamingCreate) { AddInput(InputObject); return; }
        if (_streamDeclined) return;
        try { StreamInput(InputObject); }
        catch { AbortStreamingExport(); throw; }
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (StreamingCreate) { CompleteStreamingExport(); return; }
        if (_input.Count == 0)
        {
            throw new PSArgumentException("Provide at least one object to export.", nameof(InputObject));
        }

        if (StartRow < 1 || StartColumn < 1)
        {
            throw new PSArgumentException("StartRow and StartColumn must be 1 or greater.");
        }

        if (AppendsToNamedTable && NoTable.IsPresent)
        {
            throw new PSArgumentException("TableName cannot be combined with NoTable when appending.");
        }

        var style = NoTableStyle ? ExcelTableStyle.None : TableStyle;

        var columnFormatPlan = BuildColumnFormatPlan();

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (File.Exists(resolvedPath) && NoClobber.IsPresent && !Append.IsPresent && !ReplaceSheet.IsPresent)
        {
            throw new IOException($"File '{resolvedPath}' already exists.");
        }

        var preserveWorkbook = File.Exists(resolvedPath) && (Append.IsPresent || ReplaceSheet.IsPresent);
        var action = preserveWorkbook
            ? Append.IsPresent ? "Update existing Excel workbook" : "Replace worksheet in existing Excel workbook"
            : File.Exists(resolvedPath) ? "Overwrite Excel workbook" : "Write new Excel workbook";
        if (!ShouldProcess(resolvedPath, action))
        {
            return;
        }

        var saveOptions = ExcelDocumentService.NativeSaveOptions(SaveOptions, null, null, false);

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

        var document = preserveWorkbook
            ? ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false)
            : ExcelDocumentService.CreateDocument(resolvedPath, autoSave: false);

        try
        {
            if (DateSystem.HasValue) document.DateSystem = DateSystem.Value;
            if (Theme != null) document.ApplyTheme(Theme.Resolve(OfficeTheme.FromDocument));
            if (Sheet != null && (Sheet.Position.HasValue || Sheet.Sheet != null)) {
                if (!preserveWorkbook && Sheet.Position == 1) _resolvedSheetName = "Sheet1";
                else _resolvedSheetName = Sheet.Resolve(document.Sheets, s => s.Name).Name;
            }
            if (DocumentProperty != null)
                foreach (DictionaryEntry property in DocumentProperty)
                    ExcelDocumentPropertyService.SetProperty(document, Convert.ToString(property.Key, CultureInfo.InvariantCulture)!, property.Value);

            var dataSet = ExcelTabularInputService.TryGetSingleDataSet(_input);
            if (dataSet != null)
            {
                ExportDataSet(document, dataSet, style, columnFormatPlan);
                ExcelDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, saveOptions: saveOptions);
                WritePassThru(resolvedPath);
                return;
            }

            var isAppendingToExistingSheet = Append.IsPresent && SheetExists(document, WorksheetName);
            if (AppendsToNamedTable && !isAppendingToExistingSheet)
            {
                throw new PSArgumentException($"Worksheet '{WorksheetName}' must exist when appending to a named table.");
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
                sheet.Cell(dataStartRow, StartColumn).SetValue(Title!);
                sheet.Cell(dataStartRow, StartColumn).SetBold(true);
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
                sheet.View.Freeze(frozenRows, frozenColumns);
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

}
