using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static VisioStyleTheme ResolveNativeTheme(VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options, OfficeVisioVisualConversionReport report) {
        if (options.NativeTheme != null) return options.NativeTheme.Clone();
        var theme = VisioStyleTheme.Minimal();
        var source = envelope.Presentation?.Theme;
        var foreground = source == null ? null : MapNativeColor(source.Foreground, "Theme foreground", envelope.Id, report, OfficeVisioVisualEntityKind.Artifact);
        var border = source == null ? null : MapNativeColor(source.Border, "Theme border", envelope.Id, report, OfficeVisioVisualEntityKind.Artifact);
        var connectorColor = source == null ? null : MapNativeColor(source.MutedForeground, "Theme connector", envelope.Id, report, OfficeVisioVisualEntityKind.Artifact);
        var card = source == null ? null : MapNativeColor(source.Card, "Theme card", envelope.Id, report, OfficeVisioVisualEntityKind.Artifact);
        var surface = source == null ? null : MapNativeColor(source.Surface, "Theme surface", envelope.Id, report, OfficeVisioVisualEntityKind.Artifact);
        foreach (var style in new[] { theme.Primary, theme.Success, theme.Decision, theme.Marker, theme.Emphasis, theme.Container }) {
            if (border.HasValue) style.LineColor = border.Value;
            var fill = ReferenceEquals(style, theme.Container) ? surface : card;
            if (fill.HasValue) style.FillColor = fill.Value;
            style.TextStyle ??= new VisioTextStyle();
            style.TextStyle.FontFamily = "Arial";
            style.TextStyle.Bold = false;
            if (foreground.HasValue) style.TextStyle.Color = foreground.Value;
        }
        foreach (var style in new[] { theme.Connector, theme.DataConnector, theme.ControlConnector }) {
            if (connectorColor.HasValue) style.LineColor = connectorColor.Value;
            style.TextStyle ??= new VisioTextStyle();
            style.TextStyle.FontFamily = "Arial";
            if (foreground.HasValue) style.TextStyle.Color = foreground.Value;
        }
        theme.TitleText.FontFamily = "Arial";
        theme.LegendText.FontFamily = "Arial";
        if (foreground.HasValue) {
            theme.TitleText.Color = foreground.Value;
            theme.LegendText.Color = foreground.Value;
        }
        return theme;
    }
}
