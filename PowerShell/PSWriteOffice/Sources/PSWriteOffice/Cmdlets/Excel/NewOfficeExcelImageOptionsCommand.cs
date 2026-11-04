using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates Excel rendering settings for Export-OfficeDocumentImage and ConvertTo-OfficeDocumentImage.</summary>
/// <example>
/// <summary>Render a sheet with gridlines</summary>
/// <code>$image = New-OfficeExcelImageOptions -ShowGridlines
/// Export-OfficeDocumentImage -Path .\Reports\Q3-Sales.xlsx -Sheet Summary -OutputPath .\Images\Q3-Summary.png -Options $image</code>
/// <para>Renders the Summary sheet to PNG with gridlines shown.</para>
/// </example>
/// <example>
/// <summary>Render dashboard sheets for a status page</summary>
/// <code>$image = New-OfficeExcelImageOptions -IncludeCharts -IncludeImages -IncludeConditionalFormatting
/// Export-OfficeDocumentImage -Path .\Reports\Operations.xlsx -Sheet Dashboard -OutputDirectory .\Site\Images -Scale 2 -Options $image</code>
/// <para>Renders the dashboard at double scale with charts and images, so the PNG stays sharp when embedded in a web page.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelImageOptions", DefaultParameterSetName = "Default")]
public sealed class NewOfficeExcelImageOptionsCommand : PSCmdlet {
    /// <summary>Renders worksheet gridlines; explicit false hides them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowGridlines { get; set; }

    /// <summary>Renders hidden rows and columns; explicit false omits them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeHidden { get; set; }

    /// <summary>Includes embedded worksheet images in the rendered output.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeImages { get; set; }

    /// <summary>Includes charts in the rendered output; explicit false omits them.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeCharts { get; set; }

    /// <summary>Includes worksheet drawing objects in the rendered output.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeDrawingObjects { get; set; }

    /// <summary>Evaluates supported conditional formatting during rendering; explicit false disables it.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeConditionalFormatting { get; set; }

    /// <summary>Renders hyperlink hints in worksheet snapshots.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowHyperlinkHints { get; set; }

    /// <summary>Renders comment text rather than only comment indicators.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowCommentBodies { get; set; }

    /// <summary>Maximum worksheet cells materialized for one visual snapshot.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaximumRenderedCells { get; set; }

    /// <summary>Maximum output width in pixels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaximumOutputWidth { get; set; }

    /// <summary>Maximum output height in pixels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaximumOutputHeight { get; set; }

    /// <summary>Maximum pixel count allocated for one raster image.</summary>
    [Parameter(ParameterSetName = "Default")]
    public long? MaximumRasterPixelCount { get; set; }

    /// <summary>Native policy when the requested raster exceeds its pixel budget.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeRasterOverflowBehavior? RasterOverflowBehavior { get; set; }

    /// <summary>Maximum number of images produced by one batch operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaximumOutputCount { get; set; }

    /// <summary>Maximum aggregate raster pixels in one batch.</summary>
    [Parameter(ParameterSetName = "Default")]
    public long? MaximumTotalRasterPixelCount { get; set; }

    /// <summary>Maximum aggregate encoded image bytes in one batch.</summary>
    [Parameter(ParameterSetName = "Default")]
    public long? MaximumTotalEncodedBytes { get; set; }

    /// <summary>Render timeout in seconds; omission retains the native timeout.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? RenderTimeoutSeconds { get; set; }

    /// <summary>Maximum concurrent render workers; omission retains the native limit.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaximumDegreeOfParallelism { get; set; }

    /// <summary>Language tag used for native text shaping.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? TextShapingLanguage { get; set; }

    /// <summary>Includes hidden worksheets when no explicit sheet selection is supplied.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeHiddenSheets { get; set; }

    /// <summary>Honor worksheet print areas.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetPrintAreas { get; set; }

    /// <summary>Creates separate images at worksheet manual page breaks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SplitWorksheetsByManualPageBreaks { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelImageExportOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShowGridlines))) options.ShowGridlines = ShowGridlines;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeHidden))) options.IncludeHidden = IncludeHidden;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeImages))) options.IncludeImages = IncludeImages;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeCharts))) options.IncludeCharts = IncludeCharts;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeDrawingObjects))) options.IncludeDrawingObjects = IncludeDrawingObjects;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeConditionalFormatting))) options.IncludeConditionalFormatting = IncludeConditionalFormatting;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShowHyperlinkHints))) options.ShowHyperlinkHints = ShowHyperlinkHints;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShowCommentBodies))) options.ShowCommentBodies = ShowCommentBodies;
        if (MaximumRenderedCells.HasValue) options.MaximumRenderedCells = MaximumRenderedCells.Value;
        if (MaximumOutputWidth.HasValue) options.MaximumOutputWidth = MaximumOutputWidth.Value;
        if (MaximumOutputHeight.HasValue) options.MaximumOutputHeight = MaximumOutputHeight.Value;
        if (MaximumRasterPixelCount.HasValue) options.MaximumRasterPixels = MaximumRasterPixelCount.Value;
        if (RasterOverflowBehavior.HasValue) options.RasterOverflowBehavior = RasterOverflowBehavior.Value;
        if (MaximumOutputCount.HasValue) options.MaximumOutputCount = MaximumOutputCount.Value;
        if (MaximumTotalRasterPixelCount.HasValue) options.MaximumTotalRasterPixels = MaximumTotalRasterPixelCount.Value;
        if (MaximumTotalEncodedBytes.HasValue) options.MaximumTotalEncodedBytes = MaximumTotalEncodedBytes.Value;
        if (RenderTimeoutSeconds.HasValue) options.RenderTimeout = TimeSpan.FromSeconds(RenderTimeoutSeconds.Value);
        if (MaximumDegreeOfParallelism.HasValue) options.MaximumDegreeOfParallelism = MaximumDegreeOfParallelism.Value;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TextShapingLanguage))) options.TextShapingLanguage = TextShapingLanguage;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeHiddenSheets))) options.IncludeHiddenSheets = IncludeHiddenSheets;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(UseWorksheetPrintAreas))) options.UsePrintArea = UseWorksheetPrintAreas;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(SplitWorksheetsByManualPageBreaks))) options.SplitByManualPageBreaks = SplitWorksheetsByManualPageBreaks;
        WriteObject(options);
    }

}
