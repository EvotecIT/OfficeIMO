using System;
using System.Collections.Generic;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.ChartForgeX;

static partial class OfficeVisioVisualConversionExtensions {
    private static int? MapGraphLinePattern(
        string? lineStyle,
        string edgeId,
        OfficeVisioVisualConversionReport report) {
        if (string.IsNullOrWhiteSpace(lineStyle) || string.Equals(lineStyle, "Auto", StringComparison.OrdinalIgnoreCase)) return null;
        if (string.Equals(lineStyle, "Solid", StringComparison.OrdinalIgnoreCase)) return 1;
        if (string.Equals(lineStyle, "Dashed", StringComparison.OrdinalIgnoreCase)) return 2;
        if (string.Equals(lineStyle, "Dotted", StringComparison.OrdinalIgnoreCase)) return 3;
        report.Warn($"Edge '{edgeId}' line style '{lineStyle}' was normalized to the native Visio theme pattern.");
        return null;
    }

    private static int? MapSequenceLinePattern(
        string? lineStyle,
        string messageId,
        OfficeVisioVisualConversionReport report) {
        if (string.IsNullOrWhiteSpace(lineStyle) || string.Equals(lineStyle, "Auto", StringComparison.OrdinalIgnoreCase)) return null;
        if (string.Equals(lineStyle, "Solid", StringComparison.OrdinalIgnoreCase)) return 1;
        if (string.Equals(lineStyle, "Dashed", StringComparison.OrdinalIgnoreCase)) return 2;
        if (string.Equals(lineStyle, "Dotted", StringComparison.OrdinalIgnoreCase)) return 3;
        report.Warn($"Sequence message '{messageId}' line style '{lineStyle}' was normalized to the native Visio theme pattern.");
        return null;
    }

    private static Color? MapNativeColor(
        string? value,
        string entityKind,
        string entityId,
        OfficeVisioVisualConversionReport report) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (Color.TryParseCss(value, out Color color)) return color;
        report.Warn($"{entityKind} '{entityId}' color '{value}' remains in the CFX envelope and, when enabled, Shape Data because it is not a supported native Visio color token.");
        return null;
    }

    private static void PreserveTooltipFidelity(
        IDictionary<string, string?> shapeData,
        string? tooltip,
        string? href,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        string entityKind,
        string entityId) {
        if (string.IsNullOrWhiteSpace(tooltip) || options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(href)) return;
        if (options.IncludeShapeData) {
            AddValue(shapeData, "CFX.Tooltip", tooltip);
            string reason = options.IncludeHyperlinks
                ? "native hyperlink descriptions require a hyperlink address"
                : "native hyperlink projection was disabled";
            report.Warn($"{entityKind} '{entityId}' tooltip was retained as Shape Data because {reason}.");
            return;
        }

        report.Warn($"{entityKind} '{entityId}' tooltip remains only in the CFX envelope because it could not be attached to a native hyperlink and Shape Data projection was disabled.");
    }
}
