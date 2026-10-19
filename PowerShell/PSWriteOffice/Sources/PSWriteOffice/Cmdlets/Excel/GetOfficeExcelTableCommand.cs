using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets Excel tables defined in a workbook.</summary>
/// <example>
///   <summary>List tables and export table metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$tables = Get-OfficeExcelTable -Path .\report.xlsx -Sheet Data
/// $tables |
///     Select-Object -Property Name, Sheet, Range |
///     Export-Csv -Path .\ExcelTables.csv -NoTypeInformation</code>
///   <para>Returns table metadata for workbook documentation or generated-artifact proof.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTable", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelTableCommand : AsyncPSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Remote workbook URI to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Optional table name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional sheet name filter.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional sheet index (0-based) filter.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        ExcelDocument? document = null;
        var dispose = false;

        try {
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                if (!File.Exists(resolvedPath)) {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }
                document = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            } else if (ParameterSetName == ParameterSetUri) {
                if (Uri == null) {
                    throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
                }

                document = await ExcelDocumentService.LoadDocumentAsync(
                    Uri,
                    readOnly: true,
                    allowHttp: AllowHttp.IsPresent,
                    cancellationToken: CancelToken).ConfigureAwait(false);
                dispose = true;
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Excel workbook was not provided.");
            }

            var sheetFilter = ResolveSheetName(document);
            var tables = document.GetTables();

            foreach (var table in tables) {
                if (!string.IsNullOrWhiteSpace(sheetFilter) &&
                    !string.Equals(table.SheetName, sheetFilter, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(Name) &&
                    !string.Equals(table.Name, Name, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                WriteObject(CreateRecord(
                    table.Name,
                    table.Range,
                    table.SheetName,
                    ParameterSetName == ParameterSetPath ? Path : null,
                    ParameterSetName == ParameterSetUri ? Uri : null));
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private string? ResolveSheetName(ExcelDocument document) {
        if (!string.IsNullOrWhiteSpace(Sheet)) {
            return Sheet;
        }

        if (SheetIndex.HasValue) {
            if (SheetIndex.Value < 0 || SheetIndex.Value >= document.Sheets.Count) {
                throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
            }
            return document.Sheets[SheetIndex.Value].Name;
        }

        return null;
    }

    private static PSObject CreateRecord(string name, string range, string sheet, string? path, Uri? uri) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Name", name));
        record.Properties.Add(new PSNoteProperty("Range", range));
        record.Properties.Add(new PSNoteProperty("Sheet", sheet));
        record.Properties.Add(new PSNoteProperty("WorksheetName", sheet));
        if (!string.IsNullOrWhiteSpace(path)) {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("Path", path));
        }
        if (uri != null) {
            record.Properties.Add(new PSNoteProperty("Uri", uri));
        }
        return record;
    }
}