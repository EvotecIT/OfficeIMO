using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private (double Scale, OfficeChartBubbleSizeMode Mode, bool ShowNegative, IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics)
        ReadBubbleImageSettings(C.PlotArea? plot, ExcelChartData data) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var bubble = plot?.GetFirstChild<C.BubbleChart>();
        double scale = 100D;
        var mode = OfficeChartBubbleSizeMode.Area;
        bool showNegative = false;
        if (bubble != null) {
            string? rawScale = bubble.GetFirstChild<C.BubbleScale>()?.GetAttribute("val", string.Empty).Value;
            if (rawScale != null && (!double.TryParse(rawScale, NumberStyles.Integer, CultureInfo.InvariantCulture, out scale) || scale < 0D || scale > 300D)) {
                scale = 100D;
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated,
                    "Invalid bubbleScale is rendered at 100 percent; supported values are integers from 0 through 300.");
            }
            string? representation = bubble.GetFirstChild<C.SizeRepresents>()?.GetAttribute("val", string.Empty).Value;
            if (representation == "w") mode = OfficeChartBubbleSizeMode.Width;
            else if (representation != null && representation != "area")
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated, "Unsupported sizeRepresents value is rendered as bubble area.");
            showNegative = bubble.GetFirstChild<C.ShowNegativeBubbles>() is C.ShowNegativeBubbles negative && negative.Val?.Value != false;
            if (showNegative && data.Series.Any(series => series.BubbleSizes?.Any(size => size < 0D) == true))
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported,
                    "showNegBubbles requests negative-size bubbles, which the shared renderer cannot paint; negative-size points are omitted. Positive bubble sizes are preserved.");
            if (bubble.Descendants<C.Bubble3D>().Any(item => item.Val?.Value != false))
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated, "Three-dimensional bubbles are rendered as flat bubbles.");
        }
        return (scale, mode, showNegative, diagnostics);

        void Warn(string code, string message) => diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
            OfficeImageExportDiagnosticSeverity.Warning, code, message, GetImageExportSource()));
    }
}
