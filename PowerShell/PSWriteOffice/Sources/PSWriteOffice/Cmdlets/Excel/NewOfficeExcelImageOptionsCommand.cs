using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates discoverable rendering settings for Excel range and chart image export.</summary>
/// <example>
///   <summary>Render a range with gridlines and hyperlinks visible.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelImageOptions -ShowGridlines -ShowHyperlinkHints -TargetDpi 144
/// Export-OfficeExcelRangeImage -Path .\Workbook.xlsx -Worksheet Summary -Range A1:H20 -OutputPath .\Summary.svg -Options $options</code>
/// </example>
/// <example>
///   <summary>Reuse the same rendering controls for a chart.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelImageOptions -TargetDpi 144 -MaximumOutputWidth 1600
/// Export-OfficeExcelChartImage -Path .\Workbook.xlsx -Worksheet Summary -ChartName Revenue -OutputPath .\Revenue.svg -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelImageOptions")]
[OutputType(typeof(ExcelImageExportOptions))]
public sealed class NewOfficeExcelImageOptionsCommand : OfficeImageOptionsCommandBase<ExcelImageExportOptions> {
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
    /// <summary>Show hyperlink hints.</summary>
    [Parameter] public SwitchParameter ShowHyperlinkHints { get; set; }
    /// <summary>Show cell comment bodies.</summary>
    [Parameter] public SwitchParameter ShowCommentBodies { get; set; }
    /// <summary>Maximum cells rendered.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumRenderedCells { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelImageExportOptions();
        ApplyExcel(options);
        WriteObject(options);
    }

    internal void ApplyExcel(ExcelImageExportOptions options) {
        ApplyCommon(options);
        if (IsBound(nameof(ShowGridlines))) options.ShowGridlines = ShowGridlines.IsPresent;
        if (IsBound(nameof(IncludeHidden))) options.IncludeHidden = IncludeHidden.IsPresent;
        if (IsBound(nameof(IncludeImages))) options.IncludeImages = IncludeImages.IsPresent;
        if (IsBound(nameof(IncludeCharts))) options.IncludeCharts = IncludeCharts.IsPresent;
        if (IsBound(nameof(IncludeDrawingObjects))) options.IncludeDrawingObjects = IncludeDrawingObjects.IsPresent;
        if (IsBound(nameof(IncludeConditionalFormatting))) options.IncludeConditionalFormatting = IncludeConditionalFormatting.IsPresent;
        if (IsBound(nameof(ShowHyperlinkHints))) options.ShowHyperlinkHints = ShowHyperlinkHints.IsPresent;
        if (IsBound(nameof(ShowCommentBodies))) options.ShowCommentBodies = ShowCommentBodies.IsPresent;
        if (MaximumRenderedCells.HasValue) options.MaximumRenderedCells = MaximumRenderedCells.Value;
    }
}
