using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets defined names (named ranges) from an Excel workbook.</summary>
/// <example>
///   <summary>List workbook and sheet-scoped names.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ranges = Get-OfficeExcelNamedRange -Path .\report.xlsx
/// $ranges |
///     Select-Object -Property Name, Scope, Range |
///     Format-Table -AutoSize</code>
///   <para>Returns named ranges so workbook navigation and formulas can be reviewed.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelNamedRange", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelNamedRangeCommand : AsyncPSCmdlet {
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

    /// <summary>Optional named range to retrieve.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional sheet name for sheet-scoped names.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional sheet index (0-based) for sheet-scoped names.</summary>
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

            var scope = ResolveSheet(document);

            if (!string.IsNullOrWhiteSpace(Name)) {
                var range = document.GetNamedRange(Name!, scope);
                if (range != null) {
                    WriteObject(CreateRecord(
                        Name!,
                        range,
                        scope,
                        ParameterSetName == ParameterSetPath ? Path : null,
                        ParameterSetName == ParameterSetUri ? Uri : null));
                }
                return;
            }

            var ranges = document.GetAllNamedRanges(scope);
            foreach (var entry in ranges) {
                WriteObject(CreateRecord(
                    entry.Key,
                    entry.Value,
                    scope,
                    ParameterSetName == ParameterSetPath ? Path : null,
                    ParameterSetName == ParameterSetUri ? Uri : null));
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private ExcelSheet? ResolveSheet(ExcelDocument document) {
        if (!string.IsNullOrWhiteSpace(Sheet)) {
            return document[Sheet!];
        }

        if (SheetIndex.HasValue) {
            if (SheetIndex.Value < 0 || SheetIndex.Value >= document.Sheets.Count) {
                throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
            }
            return document.Sheets[SheetIndex.Value];
        }

        return null;
    }

    private static PSObject CreateRecord(string name, string range, ExcelSheet? scope, string? path, Uri? uri) {
        var record = new PSObject();
        var sheetName = scope?.Name;
        record.Properties.Add(new PSNoteProperty("Name", name));
        record.Properties.Add(new PSNoteProperty("Range", NormalizeRange(range)));
        record.Properties.Add(new PSNoteProperty("Scope", sheetName ?? "Workbook"));
        if (!string.IsNullOrWhiteSpace(sheetName)) {
            record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
            record.Properties.Add(new PSNoteProperty("WorksheetName", sheetName));
        }
        if (!string.IsNullOrWhiteSpace(path)) {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("Path", path));
        }
        if (uri != null) {
            record.Properties.Add(new PSNoteProperty("Uri", uri));
        }
        return record;
    }

    private static string NormalizeRange(string range) {
        if (string.IsNullOrWhiteSpace(range)) {
            return range;
        }

        var separatorIndex = FindSheetSeparator(range);
        if (separatorIndex >= 0) {
            var prefix = range.Substring(0, separatorIndex + 1);
            var reference = range.Substring(separatorIndex + 1);
            return prefix + NormalizeA1Reference(reference);
        }

        return NormalizeA1Reference(range);
    }

    private static string NormalizeA1Reference(string reference) {
        return Regex.IsMatch(reference, @"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$")
            ? reference.Replace("$", string.Empty)
            : reference;
    }

    private static int FindSheetSeparator(string range) {
        var inQuotedSheetName = false;
        for (var i = 0; i < range.Length; i++) {
            if (range[i] == '\'') {
                if (inQuotedSheetName && i + 1 < range.Length && range[i + 1] == '\'') {
                    i++;
                    continue;
                }

                inQuotedSheetName = !inQuotedSheetName;
                continue;
            }

            if (!inQuotedSheetName && range[i] == '!') {
                return i;
            }
        }

        return -1;
    }
}