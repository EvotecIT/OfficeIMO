using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates discoverable sheet selection and rendering settings for Export-OfficeExcelImage.</summary>
/// <example>
///   <summary>Render selected worksheets with charts and conditional formatting.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelWorkbookImageOptions -SheetName Summary,Data -IncludeCharts -IncludeConditionalFormatting
/// Export-OfficeExcelImage -Path .\Workbook.xlsx -OutputPath .\Sheets -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelWorkbookImageOptions")]
[OutputType(typeof(ExcelWorkbookImageExportOptions))]
public sealed class NewOfficeExcelWorkbookImageOptionsCommand : OfficeImageOptionsCommandBase<ExcelWorkbookImageExportOptions> {
    /// <summary>Worksheet names to export.</summary>
    [Parameter] public string[]? SheetName { get; set; }
    /// <summary>Include hidden worksheets.</summary>
    [Parameter] public SwitchParameter IncludeHiddenSheets { get; set; }
    /// <summary>Use worksheet print areas.</summary>
    [Parameter] public SwitchParameter UseWorksheetPrintAreas { get; set; }
    /// <summary>Split worksheets at manual page breaks.</summary>
    [Parameter] public SwitchParameter SplitWorksheetsByManualPageBreaks { get; set; }
    /// <summary>Show worksheet gridlines.</summary>
    [Parameter] public SwitchParameter ShowGridlines { get; set; }
    /// <summary>Include hidden rows and columns.</summary>
    [Parameter] public SwitchParameter IncludeHidden { get; set; }
    /// <summary>Include worksheet images.</summary>
    [Parameter] public SwitchParameter IncludeImages { get; set; }
    /// <summary>Include worksheet charts.</summary>
    [Parameter] public SwitchParameter IncludeCharts { get; set; }
    /// <summary>Include drawing objects.</summary>
    [Parameter] public SwitchParameter IncludeDrawingObjects { get; set; }
    /// <summary>Include conditional formatting.</summary>
    [Parameter] public SwitchParameter IncludeConditionalFormatting { get; set; }
    /// <summary>Maximum cells rendered per worksheet.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumRenderedCells { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelWorkbookImageExportOptions();
        ApplyCommon(options);
        if (SheetName != null) options.SheetNames = SheetName;
        if (IsBound(nameof(IncludeHiddenSheets))) options.IncludeHiddenSheets = IncludeHiddenSheets.IsPresent;
        if (IsBound(nameof(UseWorksheetPrintAreas))) options.UseWorksheetPrintAreas = UseWorksheetPrintAreas.IsPresent;
        if (IsBound(nameof(SplitWorksheetsByManualPageBreaks))) options.SplitWorksheetsByManualPageBreaks = SplitWorksheetsByManualPageBreaks.IsPresent;
        if (IsBound(nameof(ShowGridlines))) options.ShowGridlines = ShowGridlines.IsPresent;
        if (IsBound(nameof(IncludeHidden))) options.IncludeHidden = IncludeHidden.IsPresent;
        if (IsBound(nameof(IncludeImages))) options.IncludeImages = IncludeImages.IsPresent;
        if (IsBound(nameof(IncludeCharts))) options.IncludeCharts = IncludeCharts.IsPresent;
        if (IsBound(nameof(IncludeDrawingObjects))) options.IncludeDrawingObjects = IncludeDrawingObjects.IsPresent;
        if (IsBound(nameof(IncludeConditionalFormatting))) options.IncludeConditionalFormatting = IncludeConditionalFormatting.IsPresent;
        if (MaximumRenderedCells.HasValue) options.MaximumRenderedCells = MaximumRenderedCells.Value;
        WriteObject(options);
    }
}
