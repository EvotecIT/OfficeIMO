using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets pivot tables defined in a workbook.</summary>
/// <example>
///   <summary>Read pivot table metadata from a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pivots = Get-OfficeExcelPivotTable -Path .\report.xlsx -Sheet Summary
/// $pivots |
///     Select-Object -Property Name, Sheet, SourceRange, RowFields, DataFields |
///     Format-List</code>
///   <para>Returns pivot table metadata so generated reports can prove the pivot source and layout.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelPivotTable", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelPivotTables")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelPivotTableCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional pivot table name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional sheet name filter.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional sheet index (0-based) filter.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
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
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Excel workbook was not provided.");
            }

            var sheetFilter = ResolveSheetName(document);
            var pivots = document.GetPivotTables();

            foreach (var pivot in pivots) {
                if (!string.IsNullOrWhiteSpace(sheetFilter) &&
                    !string.Equals(pivot.SheetName, sheetFilter, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(Name) &&
                    !string.Equals(pivot.Name, Name, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                WriteObject(CreateRecord(pivot));
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

    private static PSObject CreateRecord(ExcelPivotTableInfo pivot) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Name", pivot.Name));
        record.Properties.Add(new PSNoteProperty("Sheet", pivot.SheetName));
        record.Properties.Add(new PSNoteProperty("SheetIndex", pivot.SheetIndex));
        record.Properties.Add(new PSNoteProperty("Location", pivot.Location));
        record.Properties.Add(new PSNoteProperty("SourceSheet", pivot.SourceSheet));
        record.Properties.Add(new PSNoteProperty("SourceRange", pivot.SourceRange));
        record.Properties.Add(new PSNoteProperty("CacheId", pivot.CacheId));
        record.Properties.Add(new PSNoteProperty("PivotStyle", pivot.PivotStyle));
        record.Properties.Add(new PSNoteProperty("Layout", pivot.Layout.ToString()));
        record.Properties.Add(new PSNoteProperty("DataOnRows", pivot.DataOnRows));
        record.Properties.Add(new PSNoteProperty("ShowHeaders", pivot.ShowHeaders));
        record.Properties.Add(new PSNoteProperty("ShowEmptyRows", pivot.ShowEmptyRows));
        record.Properties.Add(new PSNoteProperty("ShowEmptyColumns", pivot.ShowEmptyColumns));
        record.Properties.Add(new PSNoteProperty("ShowDrill", pivot.ShowDrill));
        record.Properties.Add(new PSNoteProperty("RefreshOnOpen", pivot.RefreshOnOpen));
        record.Properties.Add(new PSNoteProperty("SaveSourceData", pivot.SaveSourceData));
        record.Properties.Add(new PSNoteProperty("PreserveFormatting", pivot.PreserveFormatting));
        record.Properties.Add(new PSNoteProperty("EnableDrill", pivot.EnableDrill));
        record.Properties.Add(new PSNoteProperty("RowFields", pivot.RowFields.ToArray()));
        record.Properties.Add(new PSNoteProperty("ColumnFields", pivot.ColumnFields.ToArray()));
        record.Properties.Add(new PSNoteProperty("PageFields", pivot.PageFields.ToArray()));
        record.Properties.Add(new PSNoteProperty("DataFields", CreateDataFieldRecords(pivot.DataFields)));
        return record;
    }

    private static PSObject[] CreateDataFieldRecords(IReadOnlyList<ExcelPivotDataFieldInfo> dataFields) {
        var list = new List<PSObject>(dataFields.Count);
        foreach (var field in dataFields) {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("FieldName", field.FieldName));
            record.Properties.Add(new PSNoteProperty("Function", field.Function.ToString()));
            record.Properties.Add(new PSNoteProperty("DisplayName", field.DisplayName));
            list.Add(record);
        }
        return list.ToArray();
    }
}