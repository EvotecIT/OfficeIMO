using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private (double Scale, OfficeChartBubbleSizeMode Mode, bool ShowNegative, IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics)
        ReadBubbleImageSettings(OfficeXmlNode plot, ExcelChartData data) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeXmlNode bubble = plot.Child(C, "bubbleChart");
        double scale = 100D;
        var mode = OfficeChartBubbleSizeMode.Area;
        bool showNegative = false;
        if (bubble.Exists) {
            string? rawScale = ReadImageExportBubbleVal(bubble.Child(C, "bubbleScale"));
            if (rawScale != null && (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawScale, NumberStyles.Integer, CultureInfo.InvariantCulture, out scale) || scale < 0D || scale > 300D)) {
                scale = 100D;
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated,
                    "Invalid bubbleScale is rendered at 100 percent; supported values are integers from 0 through 300.");
            }
            string? representation = ReadImageExportBubbleVal(bubble.Child(C, "sizeRepresents"));
            if (representation == "w") mode = OfficeChartBubbleSizeMode.Width;
            else if (representation != null && representation != "area")
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated, "Unsupported sizeRepresents value is rendered as bubble area.");
            OfficeXmlNode negative = bubble.Child(C, "showNegBubbles");
            showNegative = negative.Exists && negative.BoolVal() != false;
            if (showNegative && data.Series.Any(series => series.BubbleSizes?.Any(size => size < 0D) == true))
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported,
                    "showNegBubbles requests negative-size bubbles, which the shared renderer cannot paint; negative-size points are omitted. Positive bubble sizes are preserved.");
            if (bubble.Descendants(C, "bubble3D").Any(item => item.BoolVal() != false))
                Warn(ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated, "Three-dimensional bubbles are rendered as flat bubbles.");
        }
        return (scale, mode, showNegative, diagnostics);

        void Warn(string code, string message) => diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
            OfficeImageExportDiagnosticSeverity.Warning, code, message, GetImageExportSource()));
    }

    // The val attribute of a bubble setting; like the SDK's GetAttribute, an element without one cannot be read.
    private static string? ReadImageExportBubbleVal(OfficeXmlNode element) {
        if (element.IsNone) return null;
        return element.Attribute("val") ?? throw new KeyNotFoundException("The element does not allow the specified attribute.");
    }
}
