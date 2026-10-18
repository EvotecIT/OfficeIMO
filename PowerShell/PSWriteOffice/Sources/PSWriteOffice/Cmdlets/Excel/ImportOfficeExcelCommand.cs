using System;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports rows from an Excel workbook as PowerShell objects.</summary>
/// <para>Provides a fast PowerShell read command over the OfficeIMO reader pipeline.</para>
/// <example>
///   <summary>Import worksheet rows and filter pending items.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = Import-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -NumericAsDecimal
/// $rows |
///     Where-Object Status -eq 'Pending' |
///     Export-Csv -Path .\PendingRows.csv -NoTypeInformation</code>
///   <para>Reads the used range on the Data worksheet, emits PSCustomObjects, and filters them in PowerShell.</para>
/// </example>
/// <example>
///   <summary>Import every worksheet and keep the source sheet name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = Import-OfficeExcel -Path .\Workbook.xlsx -AllSheets
/// $rows | Group-Object WorksheetName</code>
///   <para>Reads the used range from each worksheet and adds a WorksheetName property to each emitted row.</para>
/// </example>
/// <example>
///   <summary>Import a worksheet by column.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeExcel -Path .\Workbook.xlsx -WorksheetName Metrics -ByColumn |
///     Where-Object ColumnName -eq 'Revenue' |
///     Select-Object -ExpandProperty Values</code>
///   <para>Returns one object per column with the column name, 1-based column index, and the column values as an array.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelImport")]
public sealed class ImportOfficeExcelCommand : AsyncPSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path to import.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "InputPath", "FullName")]
    public string? Path { get; set; }

    /// <summary>Remote workbook URI to import.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Workbook document to import from.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    [Alias("Sheet")]
    public string? WorksheetName { get; set; }

    /// <summary>Zero-based worksheet index to read.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    public int? SheetIndex { get; set; }

    /// <summary>Import all worksheets. Each emitted row or column includes WorksheetName.</summary>
    [Parameter]
    public SwitchParameter AllSheets { get; set; }

    /// <summary>Optional A1 range to read. When omitted, the used range is imported.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    public string? Range { get; set; }

    /// <summary>Starting row for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartRow { get; set; }

    /// <summary>Ending row for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndRow { get; set; }

    /// <summary>Starting column for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartColumn { get; set; }

    /// <summary>Ending column for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndColumn { get; set; }

    /// <summary>Treat all rows as data and generate column names instead of using the first row as headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Formula read mode. CachedValue returns workbook cached results; FormulaText returns formula expressions when present.</summary>
    [Parameter]
    public OfficeExcelFormulaMode FormulaMode { get; set; } = OfficeExcelFormulaMode.CachedValue;

    /// <summary>Culture used when parsing numbers and dates stored as text.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Emit rows as hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <summary>Emit a DataTable instead of enumerating row objects.</summary>
    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    /// <summary>Emit a forward-only IDataReader for database bulk-copy workflows.</summary>
    [Parameter]
    public SwitchParameter AsDataReader { get; set; }

    /// <summary>Emit one object per column with ColumnName, ColumnIndex, and Values instead of row objects.</summary>
    [Parameter]
    public SwitchParameter ByColumn { get; set; }

    /// <summary>Maximum row count inspected when -AsDataReader infers the reader schema.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int SchemaSampleSize { get; set; } = 1024;

    /// <summary>Worksheet row count requested from each streaming chunk when -AsDataReader is used.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int ChunkRows { get; set; } = 1024;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        if (!string.IsNullOrWhiteSpace(Range) && HasCoordinateRange())
        {
            throw new PSArgumentException("Specify either -Range or coordinate bounds, but not both.");
        }

        if (AsDataReader.IsPresent && (AsDataTable.IsPresent || AsHashtable.IsPresent || ByColumn.IsPresent))
        {
            throw new PSArgumentException("Specify only one of -AsDataTable, -AsDataReader, -AsHashtable, or -ByColumn.");
        }

        if (AsDataReader.IsPresent && AllSheets.IsPresent)
        {
            throw new PSArgumentException("-AsDataReader reads one worksheet range at a time. Specify a single worksheet instead of -AllSheets.");
        }

        var options = ExcelReadOutputService.CreateOptions(
            NumericAsDecimal.IsPresent,
            useCachedFormulaResult: FormulaMode == OfficeExcelFormulaMode.CachedValue,
            culture: ResolveCulture());
        options.CancellationToken = AsDataReader.IsPresent ? CancellationToken.None : CancelToken;
        options.InferSchema = AsDataReader.IsPresent || AsDataTable.IsPresent;
        options.SchemaSampleRows = SchemaSampleSize;
        options.MaxDataReaderChunkRows = ChunkRows;

        if (AllSheets.IsPresent)
        {
            if (!string.IsNullOrWhiteSpace(WorksheetName) || SheetIndex.HasValue)
            {
                throw new PSArgumentException("Specify either -AllSheets or a specific worksheet, not both.");
            }

            ExcelReadOutputService.ConfigureSelection(options, null, null, ResolveRange(), !NoHeader.IsPresent);
        }
        else
        {
            var selectedIndex = SheetIndex ?? (string.IsNullOrWhiteSpace(WorksheetName) ? 0 : null);
            ExcelReadOutputService.ConfigureSelection(options, WorksheetName, selectedIndex, ResolveRange(), !NoHeader.IsPresent);
        }

        if (AsDataReader.IsPresent)
        {
            await WriteDataReaderAsync(options).ConfigureAwait(false);
            return;
        }

        using var reader = await CreateDataReaderAsync(options).ConfigureAwait(false);
        if (AllSheets.IsPresent)
        {
            do
            {
                if (AsDataTable.IsPresent || ByColumn.IsPresent)
                {
                    var currentTable = ExcelReadOutputService.ReadCurrentResultAsDataTable(reader, reader.CurrentSheetName);
                    ExcelReadOutputService.WriteOutput(
                        this,
                        currentTable,
                        AsDataTable.IsPresent,
                        AsHashtable.IsPresent,
                        ByColumn.IsPresent,
                        reader.CurrentSheetName);
                }
                else
                {
                    ExcelReadOutputService.WriteRows(this, reader, AsHashtable.IsPresent, reader.CurrentSheetName);
                }
            }
            while (reader.NextResult());

            return;
        }

        if (AsDataTable.IsPresent || ByColumn.IsPresent)
        {
            var table = ExcelReadOutputService.ReadCurrentResultAsDataTable(reader);
            ExcelReadOutputService.WriteOutput(this, table, AsDataTable.IsPresent, AsHashtable.IsPresent, ByColumn.IsPresent, null);
        }
        else
        {
            ExcelReadOutputService.WriteRows(this, reader, AsHashtable.IsPresent, null);
        }
    }

    private async Task WriteDataReaderAsync(ExcelReadOptions options)
    {
        var reader = await CreateDataReaderAsync(options).ConfigureAwait(false);
        try
        {
            WriteObject(PSObject.AsPSObject(new OwnedDataReader(reader)), enumerateCollection: false);
        }
        catch
        {
            reader.Dispose();
            throw;
        }
    }

    private async Task<ExcelWorkbookDataReader> CreateDataReaderAsync(ExcelReadOptions options)
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Excel document was not provided.", nameof(Document));
            }

            return Document.CreateDataReader(options);
        }

        if (ParameterSetName == ParameterSetUri)
        {
            if (Uri == null)
            {
                throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
            }

            return await ExcelDocument.OpenDataReaderAsync(Uri, options, ExcelHttpLoadService.CreateOptions(AllowHttp), CancelToken)
                .ConfigureAwait(false);
        }

        if (string.IsNullOrWhiteSpace(Path))
        {
            throw new PSArgumentException("Workbook path was not provided.", nameof(Path));
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path!);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return ExcelDocument.OpenDataReader(resolvedPath, options);
    }

    private string? ResolveRange()
    {
        if (!string.IsNullOrWhiteSpace(Range))
        {
            return Range!;
        }

        if (HasCoordinateRange())
        {
            if (!StartRow.HasValue || !EndRow.HasValue || !StartColumn.HasValue || !EndColumn.HasValue)
            {
                throw new PSArgumentException("StartRow, EndRow, StartColumn, and EndColumn must all be provided when using coordinate bounds.");
            }

            if (StartRow.Value < 1 || EndRow.Value < 1 || StartColumn.Value < 1 || EndColumn.Value < 1)
            {
                throw new PSArgumentException("Coordinate bounds must be 1 or greater.");
            }

            if (StartRow.Value > EndRow.Value)
            {
                throw new PSArgumentException("StartRow must be less than or equal to EndRow.");
            }

            if (StartColumn.Value > EndColumn.Value)
            {
                throw new PSArgumentException("StartColumn must be less than or equal to EndColumn.");
            }

            return $"{A1.CellReference(StartRow.Value, StartColumn.Value)}:{A1.CellReference(EndRow.Value, EndColumn.Value)}";
        }

        return null;
    }

    private bool HasCoordinateRange()
    {
        return StartRow.HasValue || EndRow.HasValue || StartColumn.HasValue || EndColumn.HasValue;
    }

    private CultureInfo ResolveCulture()
    {
        if (string.IsNullOrWhiteSpace(CultureName))
        {
            return CultureInfo.InvariantCulture;
        }

        return CultureInfo.GetCultureInfo(CultureName!);
    }
}
