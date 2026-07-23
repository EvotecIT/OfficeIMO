using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Creates Excel image-export diagnostics with an explicit fidelity-loss classification.
    /// </summary>
    internal static class ExcelImageExportDiagnosticClassifier {
        internal static OfficeImageExportDiagnostic Create(
            OfficeImageExportDiagnosticSeverity severity,
            string code,
            string message,
            string? source = null) =>
            new OfficeImageExportDiagnostic(
                severity,
                code,
                message,
                source,
                Classify(code));

        internal static OfficeImageExportLossKind Classify(string code) {
            if (string.IsNullOrWhiteSpace(code)) {
                throw new ArgumentException("Excel image-export diagnostics require a stable code.", nameof(code));
            }

            switch (code) {
                case ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasSplit:
                case ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit:
                    return OfficeImageExportLossKind.None;

                case ExcelImageExportDiagnosticCodes.CellTextRotationApproximation:
                case ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported:
                case ExcelImageExportDiagnosticCodes.CellTextRotationUnsupported:
                case ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation:
                case ExcelImageExportDiagnosticCodes.FillPatternApproximation:
                case ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation:
                case ExcelImageExportDiagnosticCodes.ConditionalFormulaThresholdApproximation:
                case ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation:
                case ExcelImageExportDiagnosticCodes.ThreadedCommentBodyApproximation:
                case ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation:
                case ExcelImageExportDiagnosticCodes.DrawingShapeTextAutoFitUnsupported:
                case ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartKindApproximated:
                case ExcelImageExportDiagnosticCodes.ChartDataLabelPointOverridesApproximated:
                case ExcelImageExportDiagnosticCodes.ChartAreaStyleApproximation:
                case ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisStyleApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisTickLabelPositionApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisMinorTickMarkPlacementApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisCrossingApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisScaleApproximation:
                case ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation:
                case ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartSecondaryAxisUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation:
                case ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation:
                case ExcelImageExportDiagnosticCodes.PrintAreaMissing:
                case ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported:
                case ExcelImageExportDiagnosticCodes.PrintAreaUnsupported:
                case ExcelImageExportDiagnosticCodes.ManualPageBreaksSingleImageUnsupported:
                case ExcelImageExportDiagnosticCodes.PageSetupUnsupported:
                case ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted:
                case ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported:
                case ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation:
                case ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation:
                case ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation:
                    return OfficeImageExportLossKind.Approximation;

                case ExcelImageExportDiagnosticCodes.CellTextClipped:
                case ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing:
                case ExcelImageExportDiagnosticCodes.FillGradientUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded:
                case ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalColorScaleUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalDataBarUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalCellIsUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalFormulaUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalTopBottomUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalAboveAverageUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalTextRuleUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalTimePeriodUnsupported:
                case ExcelImageExportDiagnosticCodes.ConditionalDifferentialFormatUnsupported:
                case ExcelImageExportDiagnosticCodes.ImageBytesMissing:
                case ExcelImageExportDiagnosticCodes.ImageFormatUnknown:
                case ExcelImageExportDiagnosticCodes.ImageAnchorHidden:
                case ExcelImageExportDiagnosticCodes.ChartAnchorHidden:
                case ExcelImageExportDiagnosticCodes.DrawingShapeAnchorHidden:
                case ExcelImageExportDiagnosticCodes.HiddenRowsOmitted:
                case ExcelImageExportDiagnosticCodes.HiddenColumnsOmitted:
                case ExcelImageExportDiagnosticCodes.CellCommentUnsupported:
                case ExcelImageExportDiagnosticCodes.ThreadedCommentUnsupported:
                case ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartSnapshotUnavailable:
                case ExcelImageExportDiagnosticCodes.ChartKindUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartTrendlineUnsupported:
                case ExcelImageExportDiagnosticCodes.ChartDataLabelLeaderLinesUnsupported:
                case ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported:
                case ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported:
                case ExcelImageExportDiagnosticCodes.SparklineKindUnsupported:
                case ExcelImageExportDiagnosticCodes.SparklineRangeUnsupported:
                case ExcelImageExportDiagnosticCodes.SparklineExternalRangeUnsupported:
                case ExcelImageExportDiagnosticCodes.SparklineDataMissing:
                    return OfficeImageExportLossKind.Omission;

                default:
                    throw new ArgumentOutOfRangeException(
                        nameof(code),
                        code,
                        "Excel image-export diagnostic codes require an explicit fidelity-loss classification.");
            }
        }
    }
}
