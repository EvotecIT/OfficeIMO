using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets frozen panes, gridlines, direction, zoom and view mode for worksheets.</summary>
/// <example><summary>Inspect report workbooks without editing them.</summary>
/// <code>Get-ChildItem .\Reports -Filter *.xlsx | Get-OfficeExcelSheetView -Sheet Summary | Format-Table</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelSheetView", DefaultParameterSetName = "Document")]
[OutputType(typeof(ExcelSheetViewInfo))]
public sealed class GetOfficeExcelSheetViewCommand : OfficeCmdlet {
    /// <summary>Gets or sets workbook paths with filesystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets a live workbook; omission uses the current worksheet or workbook scope.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public ExcelDocument? Document { get; set; }
    /// <summary>Gets or sets worksheet names or wildcard patterns, one-based positions or native worksheet objects.</summary>
    [Parameter, SupportsWildcards, OfficeSheetTransformation, ValidateNotNullOrEmpty]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            try {
                ExcelDocument book = Document ?? OfficeDocumentContext.Find<ExcelDocument>() ?? ExcelDslContext.Current?.Document
                    ?? (Sheet?.FirstOrDefault()?.Sheet as ExcelSheet)?.Document ?? ResolveDocument(Document);
                WriteViews(book, null, useCurrentSheet: Document == null);
            } catch (ArgumentException error) { InvalidArgument(error, Document); }
            catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "excel-sheet-view-read-failed", Document);
            }
            return;
        }

        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in literal ? LiteralPath : Path) {
            IReadOnlyList<string> paths;
            try {
                paths = OfficePath.ResolveInput(SessionState, input, literal,
                    reportItemError: (error, item) => ItemError(error, "input-unreadable", item));
            } catch (ArgumentException error) { InvalidArgument(error, input); return; }
            catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "input-unreadable", input);
                continue;
            }
            foreach (string path in paths) {
                try {
                    using var book = ExcelDocumentService.LoadDocument(path, readOnly: true, autoSave: false);
                    WriteViews(book, path, useCurrentSheet: false);
                } catch (ArgumentException error) { InvalidArgument(error, input); }
                catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                    ItemError(error, "excel-sheet-view-read-failed", path);
                }
            }
        }
    }

    private void WriteViews(ExcelDocument book, string? path, bool useCurrentSheet) {
        ExcelSheet? current = useCurrentSheet && ExcelDslContext.Current?.CurrentSheet is ExcelSheet selected
            && ReferenceEquals(selected.Document, book) ? selected : null;
        foreach (ExcelSheet sheet in ExcelSheetResolver.ResolveMany(book, Sheet, current))
            WriteObject(new ExcelSheetViewInfo(sheet.Name, sheet.GetViewInfo(), path));
    }
}
