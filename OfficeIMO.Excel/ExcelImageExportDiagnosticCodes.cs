namespace OfficeIMO.Excel;

/// <summary>
/// Stable diagnostic codes emitted by Excel image export.
/// </summary>
/// <remarks>
/// These values are part of the image-export contract so callers can filter or
/// assert known unsupported and approximate rendering cases without copying
/// string literals.
/// </remarks>
public static class ExcelImageExportDiagnosticCodes {
    /// <summary>Cell text was clipped or ellipsized to fit rendered bounds.</summary>
    public const string CellTextClipped = "ExcelCellTextClipped";

    /// <summary>Cell text was suppressed because a later drawing layer covers the text anchor.</summary>
    public const string CellTextOccludedByDrawing = "ExcelCellTextOccludedByDrawing";

    /// <summary>Cell text rotation was rendered through an approximate path.</summary>
    public const string CellTextRotationApproximation = "ExcelCellTextRotationApproximation";

    /// <summary>Stacked vertical cell text is not rendered exactly yet.</summary>
    public const string CellStackedTextRotationUnsupported = "ExcelCellStackedTextRotationUnsupported";

    /// <summary>The cell requested a text rotation value outside the supported range.</summary>
    public const string CellTextRotationUnsupported = "ExcelCellTextRotationUnsupported";

    /// <summary>Rich text fell back to an approximate or plain-text layout path.</summary>
    public const string CellRichTextLayoutApproximation = "ExcelCellRichTextLayoutApproximation";

    /// <summary>Requested cell font family could not be loaded exactly by the dependency-free exporter.</summary>
    public const string CellFontFamilyFallback = "ExcelCellFontFamilyFallback";

    /// <summary>Excel gradient fills are not rendered by the dependency-free exporter yet.</summary>
    public const string FillGradientUnsupported = "ExcelFillGradientUnsupported";

    /// <summary>Excel pattern fills are rendered as deterministic hatch approximations.</summary>
    public const string FillPatternApproximation = "ExcelFillPatternApproximation";

    /// <summary>Unsupported conditional formatting rule type.</summary>
    public const string ConditionalRuleUnsupported = "ExcelConditionalRuleUnsupported";

    /// <summary>Unsupported conditional formatting icon set.</summary>
    public const string ConditionalIconSetUnsupported = "ExcelConditionalIconSetUnsupported";

    /// <summary>Conditional formatting icon sets are rendered as deterministic dependency-free approximations.</summary>
    public const string ConditionalIconSetApproximation = "ExcelConditionalIconSetApproximation";

    /// <summary>Unsupported conditional formatting color scale variant.</summary>
    public const string ConditionalColorScaleUnsupported = "ExcelConditionalColorScaleUnsupported";

    /// <summary>Unsupported conditional formatting data bar variant.</summary>
    public const string ConditionalDataBarUnsupported = "ExcelConditionalDataBarUnsupported";

    /// <summary>Unsupported conditional formatting cell-is expression.</summary>
    public const string ConditionalCellIsUnsupported = "ExcelConditionalCellIsUnsupported";

    /// <summary>Unsupported conditional formatting formula expression.</summary>
    public const string ConditionalFormulaUnsupported = "ExcelConditionalFormulaUnsupported";

    /// <summary>Unsupported conditional formatting top/bottom variant.</summary>
    public const string ConditionalTopBottomUnsupported = "ExcelConditionalTopBottomUnsupported";

    /// <summary>Unsupported conditional formatting above/below-average variant.</summary>
    public const string ConditionalAboveAverageUnsupported = "ExcelConditionalAboveAverageUnsupported";

    /// <summary>Unsupported conditional formatting text-rule variant.</summary>
    public const string ConditionalTextRuleUnsupported = "ExcelConditionalTextRuleUnsupported";

    /// <summary>Unsupported conditional formatting time-period variant.</summary>
    public const string ConditionalTimePeriodUnsupported = "ExcelConditionalTimePeriodUnsupported";

    /// <summary>Unsupported conditional formatting differential-format feature.</summary>
    public const string ConditionalDifferentialFormatUnsupported = "ExcelConditionalDifferentialFormatUnsupported";

    /// <summary>Worksheet image bytes could not be read.</summary>
    public const string ImageBytesMissing = "ExcelImageBytesMissing";

    /// <summary>Worksheet image bytes did not contain a recognized image header.</summary>
    public const string ImageFormatUnknown = "ExcelImageFormatUnknown";

    /// <summary>Worksheet image was omitted because its anchor row or column is hidden.</summary>
    public const string ImageAnchorHidden = "ExcelImageAnchorHidden";

    /// <summary>Worksheet chart was omitted because its anchor row or column is hidden.</summary>
    public const string ChartAnchorHidden = "ExcelChartAnchorHidden";

    /// <summary>Worksheet drawing shape was omitted because its anchor row or column is hidden.</summary>
    public const string DrawingShapeAnchorHidden = "ExcelDrawingShapeAnchorHidden";

    /// <summary>Hidden rows were omitted from the exported visual range.</summary>
    public const string HiddenRowsOmitted = "ExcelHiddenRowsOmitted";

    /// <summary>Hidden columns were omitted from the exported visual range.</summary>
    public const string HiddenColumnsOmitted = "ExcelHiddenColumnsOmitted";

    /// <summary>Classic comment or note bodies are not rendered yet.</summary>
    public const string CellCommentUnsupported = "ExcelCellCommentUnsupported";

    /// <summary>Classic comment or note bodies are rendered as dependency-free callout approximations.</summary>
    public const string CellCommentBodyApproximation = "ExcelCellCommentBodyApproximation";

    /// <summary>Threaded comment bodies are not rendered yet.</summary>
    public const string ThreadedCommentUnsupported = "ExcelThreadedCommentUnsupported";

    /// <summary>Threaded comment bodies are rendered as dependency-free callout approximations.</summary>
    public const string ThreadedCommentBodyApproximation = "ExcelThreadedCommentBodyApproximation";

    /// <summary>Worksheet drawing shape is not renderable by the current image exporter.</summary>
    public const string DrawingShapeUnsupported = "ExcelDrawingShapeUnsupported";

    /// <summary>Worksheet drawing shape text is rendered through an approximate rotation path.</summary>
    public const string DrawingShapeTextRotationApproximation = "ExcelDrawingShapeTextRotationApproximation";

    /// <summary>Worksheet drawing shape text requested resizing the shape to fit text, which image export does not support yet.</summary>
    public const string DrawingShapeTextAutoFitUnsupported = "ExcelDrawingShapeTextAutoFitUnsupported";

    /// <summary>Worksheet drawing shape text requested a non-horizontal orientation, which image export does not support yet.</summary>
    public const string DrawingShapeTextVerticalOrientationUnsupported = "ExcelDrawingShapeTextVerticalOrientationUnsupported";

    /// <summary>Worksheet chart could not be converted to a renderable snapshot.</summary>
    public const string ChartSnapshotUnavailable = "ExcelChartSnapshotUnavailable";

    /// <summary>Chart kind is rendered through an approximate chart snapshot.</summary>
    public const string ChartKindApproximated = "ExcelChartKindApproximated";

    /// <summary>Chart kind is not rendered by the dependency-free image exporter yet.</summary>
    public const string ChartKindUnsupported = "ExcelChartKindUnsupported";

    /// <summary>Chart trendline rendering is not supported yet.</summary>
    public const string ChartTrendlineUnsupported = "ExcelChartTrendlineUnsupported";

    /// <summary>Point-level data-label overrides are approximated.</summary>
    public const string ChartDataLabelPointOverridesApproximated = "ExcelChartDataLabelPointOverridesApproximated";

    /// <summary>Data-label leader lines are not rendered yet.</summary>
    public const string ChartDataLabelLeaderLinesUnsupported = "ExcelChartDataLabelLeaderLinesUnsupported";

    /// <summary>Chart or plot area styling is approximate.</summary>
    public const string ChartAreaStyleApproximation = "ExcelChartAreaStyleApproximation";

    /// <summary>Chart gridline styling is approximate.</summary>
    public const string ChartGridlineStyleApproximation = "ExcelChartGridlineStyleApproximation";

    /// <summary>Chart axis styling is approximate.</summary>
    public const string ChartAxisStyleApproximation = "ExcelChartAxisStyleApproximation";

    /// <summary>Chart axis tick-label placement is approximate.</summary>
    public const string ChartAxisTickLabelPositionApproximation = "ExcelChartAxisTickLabelPositionApproximation";

    /// <summary>Chart axis minor tick-mark placement is approximate.</summary>
    public const string ChartAxisMinorTickMarkPlacementApproximation = "ExcelChartAxisMinorTickMarkPlacementApproximation";

    /// <summary>Chart axis crossing behavior is approximate.</summary>
    public const string ChartAxisCrossingApproximation = "ExcelChartAxisCrossingApproximation";

    /// <summary>Chart axis scaling behavior is approximate.</summary>
    public const string ChartAxisScaleApproximation = "ExcelChartAxisScaleApproximation";

    /// <summary>Chart axis number formatting is approximate.</summary>
    public const string ChartAxisNumberFormatApproximation = "ExcelChartAxisNumberFormatApproximation";

    /// <summary>Chart secondary-axis series are rendered against the primary axis by the shared image renderer.</summary>
    public const string ChartSecondaryAxisUnsupported = "ExcelChartSecondaryAxisUnsupported";

    /// <summary>Category/date-axis number formatting is not supported yet.</summary>
    public const string ChartCategoryAxisNumberFormatUnsupported = "ExcelChartCategoryAxisNumberFormatUnsupported";

    /// <summary>Chart text styling is approximate.</summary>
    public const string ChartTextStyleApproximation = "ExcelChartTextStyleApproximation";

    /// <summary>Requested chart text font family could not be loaded exactly by the dependency-free exporter.</summary>
    public const string ChartFontFamilyFallback = "ExcelChartFontFamilyFallback";

    /// <summary>Chart series styling is approximate.</summary>
    public const string ChartSeriesStyleApproximation = "ExcelChartSeriesStyleApproximation";

    /// <summary>Print-area export was requested but the worksheet has no print area.</summary>
    public const string PrintAreaMissing = "ExcelPrintAreaMissing";

    /// <summary>Multi-area print ranges are not supported by image export yet.</summary>
    public const string PrintAreaMultipleAreasUnsupported = "ExcelPrintAreaMultipleAreasUnsupported";

    /// <summary>Multi-area print ranges were exported as separate worksheet images.</summary>
    public const string PrintAreaMultipleAreasSplit = "ExcelPrintAreaMultipleAreasSplit";

    /// <summary>Configured print area could not be used by image export.</summary>
    public const string PrintAreaUnsupported = "ExcelPrintAreaUnsupported";

    /// <summary>Manual worksheet page breaks were used to split image export into separate results.</summary>
    public const string ManualPageBreaksSplit = "ExcelManualPageBreaksSplit";

    /// <summary>Manual worksheet page-break splitting was requested through a single-image export path.</summary>
    public const string ManualPageBreaksSingleImageUnsupported = "ExcelManualPageBreaksSingleImageUnsupported";

    /// <summary>Worksheet print title rows or columns are not repeated in image page output yet.</summary>
    public const string PrintTitlesUnsupported = "ExcelPrintTitlesUnsupported";

    /// <summary>Worksheet page setup settings are not applied to image page geometry yet.</summary>
    public const string PageSetupUnsupported = "ExcelPageSetupUnsupported";

    /// <summary>Worksheet page setup image output used the default paper size because no paper size is configured.</summary>
    public const string PageSetupPaperSizeDefaulted = "ExcelPageSetupPaperSizeDefaulted";

    /// <summary>Worksheet page setup configured a paper size that image page geometry does not support yet.</summary>
    public const string PageSetupPaperSizeUnsupported = "ExcelPageSetupPaperSizeUnsupported";

    /// <summary>Worksheet headers or footers are not rendered in image page output yet.</summary>
    public const string HeaderFooterUnsupported = "ExcelHeaderFooterUnsupported";

    /// <summary>Worksheet header/footer text formatting is rendered through an approximate image-export path.</summary>
    public const string HeaderFooterFormattingApproximation = "ExcelHeaderFooterFormattingApproximation";

    /// <summary>Requested worksheet header/footer font family could not be loaded exactly by the dependency-free exporter.</summary>
    public const string HeaderFooterFontFamilyFallback = "ExcelHeaderFooterFontFamilyFallback";

    /// <summary>Worksheet header/footer image was rendered through an approximate image-export path.</summary>
    public const string HeaderFooterImageApproximation = "ExcelHeaderFooterImageApproximation";

    /// <summary>Sparkline kind is not rendered by the image exporter yet.</summary>
    public const string SparklineKindUnsupported = "ExcelSparklineKindUnsupported";

    /// <summary>Sparkline data range could not be resolved.</summary>
    public const string SparklineRangeUnsupported = "ExcelSparklineRangeUnsupported";

    /// <summary>External sparkline data ranges are not rendered yet.</summary>
    public const string SparklineExternalRangeUnsupported = "ExcelSparklineExternalRangeUnsupported";

    /// <summary>Sparkline data is missing.</summary>
    public const string SparklineDataMissing = "ExcelSparklineDataMissing";

    /// <summary>Sparklines are rendered as deterministic approximations.</summary>
    public const string SparklineRenderingApproximation = "ExcelSparklineRenderingApproximation";
}
