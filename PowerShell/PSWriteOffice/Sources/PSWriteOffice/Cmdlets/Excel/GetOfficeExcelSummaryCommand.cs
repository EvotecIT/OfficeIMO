using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets a compact structural summary of an Excel workbook.</summary>
/// <example>
///   <summary>Summarize workbook contents before release.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$summary = Get-OfficeExcelSummary -Path .\report.xlsx -IncludeSheets
/// $summary |
///     Select-Object -Property SheetCount, TableCount, ChartCount, PivotTableCount
/// $summary.Sheets |
///     Select-Object -Property Name, State, UsedRange</code>
///   <para>Returns workbook-level counts plus per-sheet tables, charts, pivots, links, comments, and used ranges.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelSummary", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelSummary")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelSummaryCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Include per-sheet details in the returned object.</summary>
    [Parameter]
    public SwitchParameter IncludeSheets { get; set; }

    /// <summary>Include OfficeIMO inspection snapshot details for schema discovery.</summary>
    [Parameter]
    public SwitchParameter IncludeSchema { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelDocument? loadedDocument = null;
        SpreadsheetDocument? spreadsheet = null;
        var dispose = false;

        try {
            ExcelDocument currentDocument;
            string? path;

            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                if (!File.Exists(resolvedPath)) {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                loadedDocument = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                currentDocument = loadedDocument;
                spreadsheet = currentDocument.OpenXmlDocument;
                path = resolvedPath;
                dispose = true;
            } else {
                if (Document == null) {
                    throw new InvalidOperationException("Excel workbook was not provided.");
                }

                currentDocument = Document;
                path = ExcelDocumentService.GetAssociatedPath(Document);
                spreadsheet = Document.OpenXmlDocument;
            }

            WriteObject(CreateSummary(spreadsheet, path, IncludeSheets.IsPresent, IncludeSchema.IsPresent, currentDocument));
        } finally {
            if (dispose) {
                loadedDocument?.Dispose();
            }
        }
    }

    private static PSObject CreateSummary(SpreadsheetDocument spreadsheet, string? path, bool includeSheets, bool includeSchema, ExcelDocument? document) {
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part was not found.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook was not found.");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
        int? activeSheetIndex = GetActiveSheetIndex(workbook, sheets.Count);
        var sheetSummaries = sheets.Select((sheet, index) => CreateSheetSummary(workbookPart, sheet, index + 1, activeSheetIndex == index)).ToArray();
        var namedRangeCount = workbook.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Count() ?? 0;

        var summary = new PSObject();
        summary.Properties.Add(new PSNoteProperty("Path", path));
        ExcelWorkbookThemeInfo? theme = document?.GetWorkbookTheme();

        summary.Properties.Add(new PSNoteProperty("DateSystem", ExcelDateSystemService.ToDisplayValue(document?.DateSystem ?? GetWorkbookDateSystem(workbook))));
        summary.Properties.Add(new PSNoteProperty("HasTheme", theme?.HasTheme ?? workbookPart.GetPartsOfType<ThemePart>().Any()));
        summary.Properties.Add(new PSNoteProperty("ThemeName", theme?.Name));
        summary.Properties.Add(new PSNoteProperty("SheetCount", sheetSummaries.Length));
        summary.Properties.Add(new PSNoteProperty("ActiveSheetIndex", activeSheetIndex));
        summary.Properties.Add(new PSNoteProperty("ActiveSheetName", activeSheetIndex.HasValue ? sheets[activeSheetIndex.Value].Name?.Value : null));
        summary.Properties.Add(new PSNoteProperty("VisibleSheetCount", sheetSummaries.Count(IsVisibleSheet)));
        summary.Properties.Add(new PSNoteProperty("HiddenSheetCount", sheetSummaries.Count(IsHiddenSheet)));
        summary.Properties.Add(new PSNoteProperty("VeryHiddenSheetCount", sheetSummaries.Count(IsVeryHiddenSheet)));
        summary.Properties.Add(new PSNoteProperty("TableCount", sheetSummaries.Sum(GetIntProperty("TableCount"))));
        summary.Properties.Add(new PSNoteProperty("ChartCount", sheetSummaries.Sum(GetIntProperty("ChartCount"))));
        summary.Properties.Add(new PSNoteProperty("PivotTableCount", sheetSummaries.Sum(GetIntProperty("PivotTableCount"))));
        summary.Properties.Add(new PSNoteProperty("SparklineGroupCount", sheetSummaries.Sum(GetIntProperty("SparklineGroupCount"))));
        summary.Properties.Add(new PSNoteProperty("SlicerPartCount", CountPivotInteractionParts(workbookPart, document, ExcelPivotInteractionCacheKind.Slicer)));
        summary.Properties.Add(new PSNoteProperty("TimelinePartCount", CountPivotInteractionParts(workbookPart, document, ExcelPivotInteractionCacheKind.Timeline)));
        summary.Properties.Add(new PSNoteProperty("ConnectionPartCount", CountPackagePartsByContentType(workbookPart, "connections")));
        summary.Properties.Add(new PSNoteProperty("QueryTablePartCount", CountPackagePartsByContentType(workbookPart, "queryTable")));
        summary.Properties.Add(new PSNoteProperty("HyperlinkCount", sheetSummaries.Sum(GetIntProperty("HyperlinkCount"))));
        summary.Properties.Add(new PSNoteProperty("CommentCount", sheetSummaries.Sum(GetIntProperty("CommentCount"))));
        summary.Properties.Add(new PSNoteProperty("NamedRangeCount", namedRangeCount));

        if (includeSheets) {
            summary.Properties.Add(new PSNoteProperty("Sheets", sheetSummaries));
        }

        if (includeSchema && document != null) {
            summary.Properties.Add(new PSNoteProperty("Schema", CreateSchemaSummary(document.CreateInspectionSnapshot())));
        }

        return summary;
    }

    private static PSObject CreateSchemaSummary(ExcelWorkbookSnapshot snapshot) {
        var schema = new PSObject();
        schema.Properties.Add(new PSNoteProperty("ActiveWorksheetIndex", snapshot.ActiveWorksheetIndex));
        schema.Properties.Add(new PSNoteProperty("ActiveWorksheetName", snapshot.ActiveWorksheetName));
        schema.Properties.Add(new PSNoteProperty("DateSystem", ExcelDateSystemService.ToDisplayValue(snapshot.DateSystem)));
        schema.Properties.Add(new PSNoteProperty("SlicerPartCount", snapshot.SlicerPartCount));
        schema.Properties.Add(new PSNoteProperty("TimelinePartCount", snapshot.TimelinePartCount));
        schema.Properties.Add(new PSNoteProperty("ConnectionPartCount", snapshot.ConnectionPartCount));
        schema.Properties.Add(new PSNoteProperty("QueryTablePartCount", snapshot.QueryTablePartCount));
        schema.Properties.Add(new PSNoteProperty("HasSlicers", snapshot.HasSlicers));
        schema.Properties.Add(new PSNoteProperty("HasTimelines", snapshot.HasTimelines));
        schema.Properties.Add(new PSNoteProperty("HasConnections", snapshot.HasConnections));
        schema.Properties.Add(new PSNoteProperty("HasQueryTables", snapshot.HasQueryTables));
        schema.Properties.Add(new PSNoteProperty("Worksheets", snapshot.Worksheets.Select(CreateSchemaWorksheet).ToArray()));
        schema.Properties.Add(new PSNoteProperty("Tables", snapshot.Worksheets.SelectMany(CreateSchemaTables).ToArray()));
        schema.Properties.Add(new PSNoteProperty("NamedRanges", snapshot.NamedRanges.Select(CreateSchemaNamedRange).ToArray()));
        schema.Properties.Add(new PSNoteProperty("FormulaCells", snapshot.Worksheets.SelectMany(CreateSchemaFormulaCells).ToArray()));
        schema.Properties.Add(new PSNoteProperty("Rows", snapshot.Worksheets.SelectMany(CreateSchemaRows).ToArray()));
        schema.Properties.Add(new PSNoteProperty("Columns", snapshot.Worksheets.SelectMany(CreateSchemaColumns).ToArray()));
        return schema;
    }

    private static PSObject CreateSchemaWorksheet(ExcelWorksheetSnapshot worksheet) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Name", worksheet.Name));
        record.Properties.Add(new PSNoteProperty("Index", worksheet.Index));
        record.Properties.Add(new PSNoteProperty("Hidden", worksheet.Hidden));
        record.Properties.Add(new PSNoteProperty("IsActive", worksheet.IsActive));
        record.Properties.Add(new PSNoteProperty("UsedRange", worksheet.UsedRangeA1));
        record.Properties.Add(new PSNoteProperty("TableCount", worksheet.Tables.Count));
        record.Properties.Add(new PSNoteProperty("FormulaCount", worksheet.Cells.Count(cell => !string.IsNullOrWhiteSpace(cell.Formula))));
        record.Properties.Add(new PSNoteProperty("DataValidationCount", worksheet.Validations.Count));
        record.Properties.Add(new PSNoteProperty("FrozenRowCount", worksheet.FrozenRowCount));
        record.Properties.Add(new PSNoteProperty("FrozenColumnCount", worksheet.FrozenColumnCount));
        record.Properties.Add(new PSNoteProperty("ShowGridlines", worksheet.ShowGridlines));
        record.Properties.Add(new PSNoteProperty("RightToLeft", worksheet.RightToLeft));
        record.Properties.Add(new PSNoteProperty("View", worksheet.View));
        record.Properties.Add(new PSNoteProperty("ZoomScale", worksheet.ZoomScale));
        record.Properties.Add(new PSNoteProperty("ZoomScaleNormal", worksheet.ZoomScaleNormal));
        record.Properties.Add(new PSNoteProperty("TabColorArgb", worksheet.TabColorArgb));
        record.Properties.Add(new PSNoteProperty("OutlineSummaryBelow", worksheet.OutlineSummaryBelow));
        record.Properties.Add(new PSNoteProperty("OutlineSummaryRight", worksheet.OutlineSummaryRight));
        record.Properties.Add(new PSNoteProperty("Protection", worksheet.Protection != null));
        return record;
    }

    private static IEnumerable<PSObject> CreateSchemaTables(ExcelWorksheetSnapshot worksheet) {
        foreach (var table in worksheet.Tables) {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("SheetName", worksheet.Name));
            record.Properties.Add(new PSNoteProperty("Name", table.Name));
            record.Properties.Add(new PSNoteProperty("Range", table.A1Range));
            record.Properties.Add(new PSNoteProperty("Style", table.StyleName));
            record.Properties.Add(new PSNoteProperty("HasHeaderRow", table.HasHeaderRow));
            record.Properties.Add(new PSNoteProperty("TotalsRowShown", table.TotalsRowShown));
            record.Properties.Add(new PSNoteProperty("Columns", table.Columns.Select(column => column.Name).ToArray()));
            yield return record;
        }
    }

    private static PSObject CreateSchemaNamedRange(ExcelNamedRangeSnapshot namedRange) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Name", namedRange.Name));
        record.Properties.Add(new PSNoteProperty("SheetName", namedRange.SheetName));
        record.Properties.Add(new PSNoteProperty("Reference", namedRange.ReferenceA1));
        record.Properties.Add(new PSNoteProperty("IsBuiltIn", namedRange.IsBuiltIn));
        return record;
    }

    private static IEnumerable<PSObject> CreateSchemaFormulaCells(ExcelWorksheetSnapshot worksheet) {
        foreach (var cell in worksheet.Cells.Where(cell => !string.IsNullOrWhiteSpace(cell.Formula))) {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("SheetName", worksheet.Name));
            record.Properties.Add(new PSNoteProperty("Address", A1.CellReference(cell.Row, cell.Column)));
            record.Properties.Add(new PSNoteProperty("Formula", cell.Formula));
            yield return record;
        }
    }

    private static IEnumerable<PSObject> CreateSchemaRows(ExcelWorksheetSnapshot worksheet) {
        foreach (var row in worksheet.Rows) {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("SheetName", worksheet.Name));
            record.Properties.Add(new PSNoteProperty("Index", row.Index));
            record.Properties.Add(new PSNoteProperty("Height", row.Height));
            record.Properties.Add(new PSNoteProperty("Hidden", row.Hidden));
            record.Properties.Add(new PSNoteProperty("CustomHeight", row.CustomHeight));
            record.Properties.Add(new PSNoteProperty("OutlineLevel", row.OutlineLevel));
            record.Properties.Add(new PSNoteProperty("Collapsed", row.Collapsed));
            yield return record;
        }
    }

    private static IEnumerable<PSObject> CreateSchemaColumns(ExcelWorksheetSnapshot worksheet) {
        foreach (var column in worksheet.Columns) {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("SheetName", worksheet.Name));
            record.Properties.Add(new PSNoteProperty("StartIndex", column.StartIndex));
            record.Properties.Add(new PSNoteProperty("EndIndex", column.EndIndex));
            record.Properties.Add(new PSNoteProperty("Width", column.Width));
            record.Properties.Add(new PSNoteProperty("Hidden", column.Hidden));
            record.Properties.Add(new PSNoteProperty("CustomWidth", column.CustomWidth));
            record.Properties.Add(new PSNoteProperty("OutlineLevel", column.OutlineLevel));
            record.Properties.Add(new PSNoteProperty("Collapsed", column.Collapsed));
            yield return record;
        }
    }

    private static PSObject CreateSheetSummary(WorkbookPart workbookPart, Sheet sheet, int index, bool isActive) {
        var state = NormalizeSheetState(sheet.State?.InnerText);
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Index", index));
        record.Properties.Add(new PSNoteProperty("Name", sheet.Name?.Value ?? string.Empty));
        record.Properties.Add(new PSNoteProperty("State", state));
        record.Properties.Add(new PSNoteProperty("IsActive", isActive));

        if (sheet.Id?.Value == null) {
            AddEmptySheetCounts(record);
            return record;
        }

        var sheetPart = workbookPart.GetPartById(sheet.Id.Value);
        if (sheetPart is ChartsheetPart chartsheetPart) {
            AddChartSheetCounts(record, chartsheetPart);
            return record;
        }

        if (sheetPart is not WorksheetPart worksheetPart) {
            AddEmptySheetCounts(record);
            return record;
        }

        var worksheet = worksheetPart.Worksheet ?? new Worksheet();
        var tableRecords = GetTableRecords(worksheetPart).ToArray();
        var chartCount = CountChartParts(worksheetPart);
        var pivotCount = worksheetPart.PivotTableParts.Count();
        var sparklineGroupCount = worksheet.Descendants<SparklineGroups>().Sum(groups => groups.Elements<SparklineGroup>().Count());
        var hyperlinkCount = worksheet.Elements<Hyperlinks>().FirstOrDefault()?.Elements<Hyperlink>().Count() ?? 0;
        var commentCount = CountComments(worksheetPart);

        record.Properties.Add(new PSNoteProperty("UsedRange", worksheet.SheetDimension?.Reference?.Value));
        record.Properties.Add(new PSNoteProperty("TableCount", tableRecords.Length));
        record.Properties.Add(new PSNoteProperty("ChartCount", chartCount));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", pivotCount));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", sparklineGroupCount));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", hyperlinkCount));
        record.Properties.Add(new PSNoteProperty("CommentCount", commentCount));
        record.Properties.Add(new PSNoteProperty("Tables", tableRecords));
        return record;
    }

    private static int? GetActiveSheetIndex(Workbook workbook, int sheetCount) {
        if (sheetCount <= 0) {
            return null;
        }

        var activeTab = workbook.GetFirstChild<BookViews>()?
            .Elements<WorkbookView>()
            .FirstOrDefault()?
            .ActiveTab?.Value ?? 0U;

        if (activeTab >= sheetCount) {
            return sheetCount - 1;
        }

        return checked((int)activeTab);
    }

    private static ExcelDateSystem GetWorkbookDateSystem(Workbook workbook)
        => workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties>()?.Date1904?.Value == true
            ? ExcelDateSystem.NineteenFour
            : ExcelDateSystem.NineteenHundred;

    private static void AddChartSheetCounts(PSObject record, ChartsheetPart chartsheetPart) {
        record.Properties.Add(new PSNoteProperty("UsedRange", null));
        record.Properties.Add(new PSNoteProperty("TableCount", 0));
        record.Properties.Add(new PSNoteProperty("ChartCount", CountChartParts(chartsheetPart)));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", 0));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", 0));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", 0));
        record.Properties.Add(new PSNoteProperty("CommentCount", 0));
        record.Properties.Add(new PSNoteProperty("Tables", Array.Empty<PSObject>()));
    }

    private static int CountChartParts(OpenXmlPartContainer container) {
        return container.Parts.Sum(part =>
            (part.OpenXmlPart is ChartPart ? 1 : 0) + CountChartParts(part.OpenXmlPart));
    }

    private static int CountPackagePartsByContentType(OpenXmlPartContainer container, string marker) {
        if (string.IsNullOrWhiteSpace(marker)) {
            return 0;
        }

        return container.Parts.Sum(part => {
            var openXmlPart = part.OpenXmlPart;
            var count = openXmlPart.ContentType.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0 ? 1 : 0;
            return count + CountPackagePartsByContentType(openXmlPart, marker);
        });
    }

    private static int CountPivotInteractionParts(
        WorkbookPart workbookPart,
        ExcelDocument? document,
        ExcelPivotInteractionCacheKind kind) {
        string marker = kind == ExcelPivotInteractionCacheKind.Slicer ? "slicer" : "timeline";
        int nativeOrLegacyCount = CountPackagePartsByContentType(workbookPart, marker);
        if (document == null) {
            return nativeOrLegacyCount;
        }

        IReadOnlyList<ExcelPivotInteractionCacheInfo> caches = kind == ExcelPivotInteractionCacheKind.Slicer
            ? document.GetWorkbookSlicerCaches()
            : document.GetWorkbookTimelineCaches();
        int combinedMetadataPartCount = caches
            .Select(cache => cache.RelationshipId)
            .Distinct(StringComparer.Ordinal)
            .Count(relationshipId => {
                OpenXmlPart part = workbookPart.GetPartById(relationshipId);
                return part.ContentType.IndexOf(marker, StringComparison.OrdinalIgnoreCase) < 0;
            });

        return nativeOrLegacyCount + combinedMetadataPartCount;
    }

    private static int CountComments(WorksheetPart worksheetPart) {
        var legacyCount = worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>().Count() ?? 0;
        var threadedCount = worksheetPart.WorksheetThreadedCommentsParts.Sum(part =>
            part.ThreadedComments?.Elements<ThreadedComment>().Count() ?? 0);
        return legacyCount + threadedCount;
    }

    private static IEnumerable<PSObject> GetTableRecords(WorksheetPart worksheetPart) {
        foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
            var table = tableDefinitionPart.Table;
            if (table == null) {
                continue;
            }

            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("Name", table.Name?.Value ?? table.DisplayName?.Value));
            record.Properties.Add(new PSNoteProperty("Range", table.Reference?.Value));
            record.Properties.Add(new PSNoteProperty("Style", table.TableStyleInfo?.Name?.Value));
            yield return record;
        }
    }

    private static void AddEmptySheetCounts(PSObject record) {
        record.Properties.Add(new PSNoteProperty("UsedRange", null));
        record.Properties.Add(new PSNoteProperty("TableCount", 0));
        record.Properties.Add(new PSNoteProperty("ChartCount", 0));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", 0));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", 0));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", 0));
        record.Properties.Add(new PSNoteProperty("CommentCount", 0));
        record.Properties.Add(new PSNoteProperty("Tables", Array.Empty<PSObject>()));
    }

    private static Func<PSObject, int> GetIntProperty(string name) {
        return record => record.Properties[name]?.Value is int value ? value : 0;
    }

    private static string NormalizeSheetState(string? state) {
        if (string.IsNullOrWhiteSpace(state)) {
            return "Visible";
        }

        return state!.Equals("veryHidden", StringComparison.OrdinalIgnoreCase)
            ? "VeryHidden"
            : state.Equals("hidden", StringComparison.OrdinalIgnoreCase)
                ? "Hidden"
                : "Visible";
    }

    private static bool IsVisibleSheet(PSObject record) {
        return string.Equals(record.Properties["State"]?.Value as string, "Visible", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsHiddenSheet(PSObject record) {
        return string.Equals(record.Properties["State"]?.Value as string, "Hidden", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsVeryHiddenSheet(PSObject record) {
        return string.Equals(record.Properties["State"]?.Value as string, "VeryHidden", StringComparison.OrdinalIgnoreCase);
    }
}