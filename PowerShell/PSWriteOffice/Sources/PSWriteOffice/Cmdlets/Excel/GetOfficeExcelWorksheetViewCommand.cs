using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets worksheet view settings such as frozen panes and gridline visibility.</summary>
/// <example>
///   <summary>Inspect worksheet view settings.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$views = Get-OfficeExcelWorksheetView -Path .\Report.xlsx
/// $views |
///     Format-Table SheetName, FrozenRowCount, FrozenColumnCount, ShowGridlines, ZoomScale</code>
///   <para>Returns view metadata useful for workbook audits and maintenance scripts.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelWorksheetView", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelWorksheetView")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelWorksheetViewCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var path = string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            WriteObject(CreateViewRecord(sheet, sheet.GetViewInfo(), path));
        }
    }

    private static PSObject CreateViewRecord(ExcelSheet sheet, ExcelWorksheetViewInfo view, string? path) {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("SheetName", sheet.Name));
        record.Properties.Add(new PSNoteProperty("Sheet", sheet.Name));
        record.Properties.Add(new PSNoteProperty("HasPane", view.HasPane));
        record.Properties.Add(new PSNoteProperty("PaneState", view.PaneState));
        record.Properties.Add(new PSNoteProperty("HorizontalSplit", view.HorizontalSplit));
        record.Properties.Add(new PSNoteProperty("VerticalSplit", view.VerticalSplit));
        record.Properties.Add(new PSNoteProperty("FrozenRowCount", view.FrozenRowCount));
        record.Properties.Add(new PSNoteProperty("FrozenColumnCount", view.FrozenColumnCount));
        record.Properties.Add(new PSNoteProperty("TopLeftCell", view.TopLeftCell));
        record.Properties.Add(new PSNoteProperty("ActivePane", view.ActivePane));
        record.Properties.Add(new PSNoteProperty("ShowGridlines", view.ShowGridlines));
        record.Properties.Add(new PSNoteProperty("RightToLeft", view.RightToLeft));
        record.Properties.Add(new PSNoteProperty("View", view.View));
        record.Properties.Add(new PSNoteProperty("ZoomScale", view.ZoomScale));
        record.Properties.Add(new PSNoteProperty("ZoomScaleNormal", view.ZoomScaleNormal));
        if (!string.IsNullOrWhiteSpace(path)) {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("Path", path));
        }

        return record;
    }
}