using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reads rows from a worksheet, range or table as objects.</summary>
/// <example>
/// <summary>Read the first sheet as objects</summary>
/// <code>Import-OfficeExcel -Path .\Reports\Q3-Sales.xlsx</code>
/// <para>Reads the used range of the first worksheet, using the header row as property names.</para>
/// </example>
/// <example>
/// <summary>Find the largest open orders</summary>
/// <code>Import-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -Sheet Orders -Range A5 |
///     Where-Object Status -eq 'Open' |
///     Sort-Object Amount -Descending |
///     Select-Object -First 10 OrderId, Customer, Region, Amount</code>
/// <para>Reads the Orders sheet from row 5, where the header starts below a title block, and lists the ten largest open orders.</para>
/// </example>
/// <example>
/// <summary>Read a named table</summary>
/// <code>Import-OfficeExcel -Path .\Reports\Helpdesk.xlsx -Table Tickets |
///     Where-Object Status -ne 'Closed' |
///     Group-Object Assignee |
///     Sort-Object Count -Descending |
///     Select-Object Name, Count</code>
/// <para>Reads the rows of the Tickets table wherever it sits in the workbook and counts open tickets per assignee.</para>
/// </example>
/// <example>
/// <summary>Combine every sheet of many workbooks</summary>
/// <code>Get-ChildItem -Path .\Branches -Filter *.xlsx |
///     Import-OfficeExcel -Sheet * |
///     Export-OfficeExcel -Path .\Reports\All-Branches.xlsx -Sheet 'All branches' -Title 'Branch sales, October 2026'</code>
/// <para>Reads every worksheet of each branch workbook (each row carries its WorksheetName) and writes one consolidated workbook.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcel", DefaultParameterSetName = "Document")]
public sealed class ImportOfficeExcelCommand : AsyncPSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "PathTable", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Uri")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>Writes data without a header row.</summary>
    [Parameter()]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Selects objects, hashtables, a DataTable, an owned IDataReader or column output.</summary>
    [Parameter()]
    public ExcelImportOutput As { get; set; } = ExcelImportOutput.Object;

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter()]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Formula read mode. CachedValue returns workbook cached results; FormulaText returns formula expressions when present.</summary>
    [Parameter()]
    public OfficeExcelFormulaMode FormulaMode { get; set; } = OfficeExcelFormulaMode.CachedValue;

    /// <summary>Culture used for value conversion and number-format presets.</summary>
    [Parameter()]
    public CultureInfo? Culture { get; set; }

    /// <summary>Maximum row count inspected when -AsDataReader infers the reader schema.</summary>
    [Parameter()]
    [ValidateRange(1, int.MaxValue)]
    public int SchemaSampleRows { get; set; } = 1024;

    /// <summary>Worksheet row count requested from each streaming chunk when -AsDataReader is used.</summary>
    [Parameter()]
    [ValidateRange(1, int.MaxValue)]
    public int ChunkRows { get; set; } = 1024;

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Uri")]
    [Parameter(ParameterSetName = "Document")]
    public string? Range { get; set; }

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "PathTable", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathTable", Mandatory = true)]
    [Parameter(ParameterSetName = "UriTable", Mandatory = true)]
    [Parameter(ParameterSetName = "DocumentTable", Mandatory = true)]
    public string Table { get; set; } = null!;

    /// <summary>Literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathTable", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Remote workbook URI to import.</summary>
    [Parameter(ParameterSetName = "Uri", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "UriTable", Mandatory = true, Position = 0)]
    public Uri Uri { get; set; } = null!;

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = "Uri")]
    [Parameter(ParameterSetName = "UriTable")]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "DocumentTable", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (!ParameterSetName.StartsWith("Document", StringComparison.Ordinal) && Sheet?.Any(selector => selector.Sheet != null) == true) throw new PSArgumentException("Native worksheet objects require their caller-owned -Document or enclosing Excel content block.");
        if (Culture != null && ExcelRuntimePreflight.InspectCurrent().GlobalizationInvariantMode) WriteWarning("Culture-aware Excel parsing is affected by globalization-invariant mode.");
        if (ParameterSetName.StartsWith("Uri", StringComparison.Ordinal)) {
            if (Table != null) {
                var document = await ExcelDocumentService.LoadDocumentAsync(Uri!, readOnly: true, allowHttp: AllowHttp, cancellationToken: CancelToken).ConfigureAwait(false);
                await ImportDocument(document, ownsDocument: true).ConfigureAwait(false);
            } else {
                var options = Options();
                options.Transport = ExcelHttpLoadService.CreateOptions(AllowHttp);
                await Emit(await ExcelDocument.OpenReaderAsync(Uri!, options, CancelToken).ConfigureAwait(false), null).ConfigureAwait(false);
            }
        } else if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal)) {
            await ImportDocument(ExcelCommandTargets.Document(Document)).ConfigureAwait(false);
        } else {
            bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
            foreach (string path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
                if (Table == null) await Emit(ExcelDocument.OpenReader(path, Options()), null).ConfigureAwait(false);
                else {
                    var document = ExcelDocumentService.LoadDocument(path, readOnly: true, autoSave: false);
                    await ImportDocument(document, ownsDocument: true).ConfigureAwait(false);
                }
            }
        }
    }
    private ExcelReaderOptions Options() => new ExcelReaderOptions {
        NumericAsDecimal = NumericAsDecimal, UseCachedFormulaResult = FormulaMode == OfficeExcelFormulaMode.CachedValue,
        Culture = Culture ?? CultureInfo.InvariantCulture, HasHeaderRow = !NoHeader, A1Range = Range,
        InferSchema = As is ExcelImportOutput.DataReader or ExcelImportOutput.DataTable,
        SchemaSampleRows = SchemaSampleRows, MaxDataReaderChunkRows = ChunkRows,
        CancellationToken = As == ExcelImportOutput.DataReader ? System.Threading.CancellationToken.None : CancelToken
    };
    private async Task ImportDocument(ExcelDocument document, bool ownsDocument = false) {
        try {
        var options = Options();
        OfficeSheetSelector[]? selection = Sheet;
        if (Table != null) {
            var matches = document.GetTables().Where(table => string.Equals(table.Name, Table, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1) throw new PSArgumentException("Table must identify exactly one table in the workbook.");
            options.SheetName = matches[0].SheetName;
            options.A1Range = matches[0].Range;
            selection = new[] { OfficeSheetSelector.From(matches[0].SheetName) };
        } else if (selection == null && Document == null && ExcelDslContext.Current?.CurrentSheet is ExcelSheet current) selection = new[] { OfficeSheetSelector.From(current) };
        if (selection != null) {
            var sheets = ExcelSheetResolver.ResolveMany(document, selection);
            selection = sheets.Select(sheet => OfficeSheetSelector.From(sheet.Name)).ToArray();
        }
        // The native reader retains its own tabular source; caller-owned workbooks are never disposed here.
        await Emit(document.CreateDataReader(options), selection, ownsDocument ? document : null).ConfigureAwait(false);
        } catch { if (ownsDocument) document.Dispose(); throw; }
    }
    private Task Emit(ExcelWorkbookDataReader reader, OfficeSheetSelector[]? selection, IDisposable? ownedDocument = null) {
        bool transferred = false;
        try {
            selection ??= Sheet;
            var selected = new HashSet<string>(OfficeSheetSelector.ResolveNames(reader.SheetNames,
                selection ?? new[] { new OfficeSheetSelector(position: 1) }), StringComparer.OrdinalIgnoreCase);
            if (As == ExcelImportOutput.DataReader && selected.Count > 1) throw new PSArgumentException("DataReader output reads one sheet at a time; choose a single sheet.");
            bool multiple = selected.Count > 1 || (Sheet ?? selection)?.Any(s => s.Name != null && WildcardPattern.ContainsWildcardCharacters(s.Name)) == true;
            do {
                if (!selected.Contains(reader.CurrentSheetName)) continue;
                if (As == ExcelImportOutput.DataReader) {
                    WriteObject(PSObject.AsPSObject(new OwnedDataReader(reader, ownedDocument)), enumerateCollection: false);
                    transferred = true;
                    break;
                }
                if (As is ExcelImportOutput.DataTable or ExcelImportOutput.Column) {
                    var data = ExcelReadOutputService.ReadCurrentResultAsDataTable(reader, reader.CurrentSheetName);
                    ExcelReadOutputService.WriteOutput(this, data, As == ExcelImportOutput.DataTable, false, As == ExcelImportOutput.Column, multiple ? reader.CurrentSheetName : null);
                } else ExcelReadOutputService.WriteRows(this, reader, As == ExcelImportOutput.Hashtable, multiple ? reader.CurrentSheetName : null);
            } while (reader.NextResult());
        } finally { if (!transferred) { try { reader.Dispose(); } finally { ownedDocument?.Dispose(); } } }
        return Task.CompletedTask;
    }

}
