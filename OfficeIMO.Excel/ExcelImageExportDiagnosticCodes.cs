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
    public const string CellTextClipped = "excel-cell-text-clipped";

    /// <summary>Cell text was suppressed because a later drawing layer covers the text anchor.</summary>
    public const string CellTextOccludedByDrawing = "excel-cell-text-occluded-by-drawing";

    /// <summary>Cell text rotation was rendered through an approximate path.</summary>
    public const string CellTextRotationApproximation = "excel-cell-text-rotation-approximation";

    /// <summary>Stacked vertical cell text is not rendered exactly yet.</summary>
    public const string CellStackedTextRotationUnsupported = "excel-cell-stacked-text-rotation-unsupported";

    /// <summary>The cell requested a text rotation value outside the supported range.</summary>
    public const string CellTextRotationUnsupported = "excel-cell-text-rotation-unsupported";

    /// <summary>Rich text fell back to an approximate or plain-text layout path.</summary>
    public const string CellRichTextLayoutApproximation = "excel-cell-rich-text-layout-approximation";

    /// <summary>Excel gradient fills are not rendered by the dependency-free exporter yet.</summary>
    public const string FillGradientUnsupported = "excel-fill-gradient-unsupported";

    /// <summary>Excel pattern fills are rendered as deterministic hatch approximations.</summary>
    public const string FillPatternApproximation = "excel-fill-pattern-approximation";

    /// <summary>Unsupported conditional formatting rule type.</summary>
    public const string ConditionalRuleUnsupported = "excel-conditional-rule-unsupported";

    /// <summary>A conditional formatting rule was omitted because its reference exceeds the bounded evaluation limit.</summary>
    public const string ConditionalReferenceLimitExceeded = "excel-conditional-reference-limit-exceeded";

    /// <summary>Unsupported conditional formatting icon set.</summary>
    public const string ConditionalIconSetUnsupported = "excel-conditional-icon-set-unsupported";

    /// <summary>Conditional formatting icon sets are rendered as deterministic dependency-free approximations.</summary>
    public const string ConditionalIconSetApproximation = "excel-conditional-icon-set-approximation";

    /// <summary>Office extension conditional-formatting semantics are projected through the dependency-free renderer.</summary>
    public const string ConditionalExtensionApproximation = "excel-conditional-extension-approximation";

    /// <summary>Unsupported conditional formatting color scale variant.</summary>
    public const string ConditionalColorScaleUnsupported = "excel-conditional-color-scale-unsupported";

    /// <summary>Unsupported conditional formatting data bar variant.</summary>
    public const string ConditionalDataBarUnsupported = "excel-conditional-data-bar-unsupported";

    /// <summary>Unsupported conditional formatting cell-is expression.</summary>
    public const string ConditionalCellIsUnsupported = "excel-conditional-cell-is-unsupported";

    /// <summary>Unsupported conditional formatting formula expression.</summary>
    public const string ConditionalFormulaUnsupported = "excel-conditional-formula-unsupported";

    /// <summary>Conditional formatting formula thresholds used deterministic fallback positions.</summary>
    public const string ConditionalFormulaThresholdApproximation = "excel-conditional-formula-threshold-approximation";

    /// <summary>Unsupported conditional formatting top/bottom variant.</summary>
    public const string ConditionalTopBottomUnsupported = "excel-conditional-top-bottom-unsupported";

    /// <summary>Unsupported conditional formatting above/below-average variant.</summary>
    public const string ConditionalAboveAverageUnsupported = "excel-conditional-above-average-unsupported";

    /// <summary>Unsupported conditional formatting text-rule variant.</summary>
    public const string ConditionalTextRuleUnsupported = "excel-conditional-text-rule-unsupported";

    /// <summary>Unsupported conditional formatting time-period variant.</summary>
    public const string ConditionalTimePeriodUnsupported = "excel-conditional-time-period-unsupported";

    /// <summary>Unsupported conditional formatting differential-format feature.</summary>
    public const string ConditionalDifferentialFormatUnsupported = "excel-conditional-differential-format-unsupported";

    /// <summary>A native picture hyperlink cannot be represented safely in interactive SVG.</summary>
    public const string ImageHyperlinkUnsupported = "excel-image-hyperlink-unsupported";

    /// <summary>Worksheet image bytes could not be read.</summary>
    public const string ImageBytesMissing = "excel-image-bytes-missing";

    /// <summary>Worksheet image bytes did not contain a recognized image header.</summary>
    public const string ImageFormatUnknown = "excel-image-format-unknown";

    /// <summary>Worksheet image was omitted because its anchor row or column is hidden.</summary>
    public const string ImageAnchorHidden = "excel-image-anchor-hidden";

    /// <summary>Worksheet chart was omitted because its anchor row or column is hidden.</summary>
    public const string ChartAnchorHidden = "excel-chart-anchor-hidden";

    /// <summary>Worksheet drawing shape was omitted because its anchor row or column is hidden.</summary>
    public const string DrawingShapeAnchorHidden = "excel-drawing-shape-anchor-hidden";

    /// <summary>Hidden rows were omitted from the exported visual range.</summary>
    public const string HiddenRowsOmitted = "excel-hidden-rows-omitted";

    /// <summary>Hidden columns were omitted from the exported visual range.</summary>
    public const string HiddenColumnsOmitted = "excel-hidden-columns-omitted";

    /// <summary>Classic comment or note bodies are not rendered yet.</summary>
    public const string CellCommentUnsupported = "excel-cell-comment-unsupported";

    /// <summary>Classic comment or note bodies are rendered as dependency-free callout approximations.</summary>
    public const string CellCommentBodyApproximation = "excel-cell-comment-body-approximation";

    /// <summary>Threaded comment bodies are not rendered yet.</summary>
    public const string ThreadedCommentUnsupported = "excel-threaded-comment-unsupported";

    /// <summary>Threaded comment bodies are rendered as dependency-free callout approximations.</summary>
    public const string ThreadedCommentBodyApproximation = "excel-threaded-comment-body-approximation";

    /// <summary>Worksheet drawing shape is not renderable by the current image exporter.</summary>
    public const string DrawingShapeUnsupported = "excel-drawing-shape-unsupported";

    /// <summary>Worksheet drawing shape text is rendered through an approximate rotation path.</summary>
    public const string DrawingShapeTextRotationApproximation = "excel-drawing-shape-text-rotation-approximation";

    /// <summary>Worksheet drawing shape text requested resizing the shape to fit text, which image export does not support yet.</summary>
    public const string DrawingShapeTextAutoFitUnsupported = "excel-drawing-shape-text-auto-fit-unsupported";

    /// <summary>Worksheet drawing shape text requested a non-horizontal orientation, which image export does not support yet.</summary>
    public const string DrawingShapeTextVerticalOrientationUnsupported = "excel-drawing-shape-text-vertical-orientation-unsupported";

    /// <summary>Worksheet chart could not be converted to a renderable snapshot.</summary>
    public const string ChartSnapshotUnavailable = "excel-chart-snapshot-unavailable";

    /// <summary>Requested negative-size bubbles cannot be painted by the shared renderer.</summary>
    public const string ChartBubbleNegativeSizesUnsupported = "excel-chart-bubble-negative-sizes-unsupported";

    /// <summary>Unsupported or invalid bubble settings use an explicitly diagnosed fallback.</summary>
    public const string ChartBubbleSettingsApproximated = "excel-chart-bubble-settings-approximated";

    /// <summary>A chart family is rendered as an explicit projection.</summary>
    public const string ChartKindApproximated = "excel-chart-kind-approximated";

    /// <summary>Chart kind is not rendered by the dependency-free image exporter yet.</summary>
    public const string ChartKindUnsupported = "excel-chart-kind-unsupported";

    /// <summary>Chart axis units would exceed the shared renderer tick budget.</summary>
    public const string ChartAxisUnitUnsupported = "excel-chart-axis-unit-unsupported";

    /// <summary>Chart trendline rendering is not supported yet.</summary>
    public const string ChartTrendlineUnsupported = "excel-chart-trendline-unsupported";

    /// <summary>Point-level data-label overrides are approximated.</summary>
    public const string ChartDataLabelPointOverridesApproximated = "excel-chart-data-label-point-overrides-approximated";

    /// <summary>Data-label leader lines are not rendered yet.</summary>
    public const string ChartDataLabelLeaderLinesUnsupported = "excel-chart-data-label-leader-lines-unsupported";

    /// <summary>Chart or plot area styling is approximate.</summary>
    public const string ChartAreaStyleApproximation = "excel-chart-area-style-approximation";

    /// <summary>Chart gridline styling is approximate.</summary>
    public const string ChartGridlineStyleApproximation = "excel-chart-gridline-style-approximation";

    /// <summary>Chart axis styling is approximate.</summary>
    public const string ChartAxisStyleApproximation = "excel-chart-axis-style-approximation";

    /// <summary>Chart axis tick-label placement is approximate.</summary>
    public const string ChartAxisTickLabelPositionApproximation = "excel-chart-axis-tick-label-position-approximation";

    /// <summary>Chart axis minor tick-mark placement is approximate.</summary>
    public const string ChartAxisMinorTickMarkPlacementApproximation = "excel-chart-axis-minor-tick-mark-placement-approximation";

    /// <summary>Chart axis crossing behavior is approximate.</summary>
    public const string ChartAxisCrossingApproximation = "office-chart-axis-crossing-position-approximation";

    /// <summary>Chart axis scaling behavior is approximate.</summary>
    public const string ChartAxisScaleApproximation = "excel-chart-axis-scale-approximation";

    /// <summary>Chart axis number formatting is approximate.</summary>
    public const string ChartAxisNumberFormatApproximation = "excel-chart-axis-number-format-approximation";

    /// <summary>Chart secondary-axis series are rendered against the primary axis by the shared image renderer.</summary>
    public const string ChartSecondaryAxisUnsupported = "excel-chart-secondary-axis-unsupported";

    /// <summary>Category/date-axis number formatting is not supported yet.</summary>
    public const string ChartCategoryAxisNumberFormatUnsupported = "excel-chart-category-axis-number-format-unsupported";

    /// <summary>Chart text styling is approximate.</summary>
    public const string ChartTextStyleApproximation = "excel-chart-text-style-approximation";

    /// <summary>Chart series styling is approximate.</summary>
    public const string ChartSeriesStyleApproximation = "excel-chart-series-style-approximation";

    /// <summary>Print-area export was requested but the worksheet has no print area.</summary>
    public const string PrintAreaMissing = "excel-print-area-missing";

    /// <summary>Multi-area print ranges are not supported by image export yet.</summary>
    public const string PrintAreaMultipleAreasUnsupported = "excel-print-area-multiple-areas-unsupported";

    /// <summary>Multi-area print ranges were exported as separate worksheet images.</summary>
    public const string PrintAreaMultipleAreasSplit = "excel-print-area-multiple-areas-split";

    /// <summary>Configured print area could not be used by image export.</summary>
    public const string PrintAreaUnsupported = "excel-print-area-unsupported";

    /// <summary>Manual worksheet page breaks were used to split image export into separate results.</summary>
    public const string ManualPageBreaksSplit = "excel-manual-page-breaks-split";

    /// <summary>Manual worksheet page-break splitting was requested through a single-image export path.</summary>
    public const string ManualPageBreaksSingleImageUnsupported = "excel-manual-page-breaks-single-image-unsupported";

    /// <summary>Worksheet print title rows or columns are not repeated in image page output yet.</summary>
    public const string PrintTitlesUnsupported = "excel-print-titles-unsupported";

    /// <summary>Worksheet page setup settings are not applied to image page geometry yet.</summary>
    public const string PageSetupUnsupported = "excel-page-setup-unsupported";

    /// <summary>Worksheet page setup image output used the default paper size because no paper size is configured.</summary>
    public const string PageSetupPaperSizeDefaulted = "excel-page-setup-paper-size-defaulted";

    /// <summary>Worksheet page setup configured a paper size that image page geometry does not support yet.</summary>
    public const string PageSetupPaperSizeUnsupported = "excel-page-setup-paper-size-unsupported";

    /// <summary>Worksheet headers or footers are not rendered in image page output yet.</summary>
    public const string HeaderFooterUnsupported = "excel-header-footer-unsupported";

    /// <summary>Worksheet header/footer text formatting is rendered through an approximate image-export path.</summary>
    public const string HeaderFooterFormattingApproximation = "excel-header-footer-formatting-approximation";

    /// <summary>Worksheet header/footer image was rendered through an approximate image-export path.</summary>
    public const string HeaderFooterImageApproximation = "excel-header-footer-image-approximation";

    /// <summary>Sparkline kind is not rendered by the image exporter yet.</summary>
    public const string SparklineKindUnsupported = "excel-sparkline-kind-unsupported";

    /// <summary>Sparkline data range could not be resolved.</summary>
    public const string SparklineRangeUnsupported = "excel-sparkline-range-unsupported";

    /// <summary>External sparkline data ranges are not rendered yet.</summary>
    public const string SparklineExternalRangeUnsupported = "excel-sparkline-external-range-unsupported";

    /// <summary>Sparkline data is missing.</summary>
    public const string SparklineDataMissing = "excel-sparkline-data-missing";

    /// <summary>Sparklines are rendered as deterministic approximations.</summary>
    public const string SparklineRenderingApproximation = "excel-sparkline-rendering-approximation";
}
