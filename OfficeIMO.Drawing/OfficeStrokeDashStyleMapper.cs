using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Maps Office document dash and line-pattern vocabularies to the shared renderer stroke dash style.
/// </summary>
public static class OfficeStrokeDashStyleMapper {
    /// <summary>
    /// Maps a Visio <c>LinePattern</c> cell value to the shared stroke dash style.
    /// </summary>
    /// <param name="linePattern">Visio line-pattern integer value.</param>
    /// <returns>The closest shared dash style supported by OfficeIMO.Drawing renderers.</returns>
    public static OfficeStrokeDashStyle FromVisioLinePattern(int linePattern) {
        switch (linePattern) {
            case 3:
                return OfficeStrokeDashStyle.Dot;
            case 4:
                return OfficeStrokeDashStyle.DashDot;
            case 5:
                return OfficeStrokeDashStyle.DashDotDot;
            case 0:
            case 1:
                return OfficeStrokeDashStyle.Solid;
            default:
                return OfficeStrokeDashStyle.Dash;
        }
    }

    /// <summary>
    /// Maps an Office preset dash value name to the shared stroke dash style.
    /// </summary>
    /// <param name="presetDash">Preset dash value as an enum name or serialized token.</param>
    /// <param name="dashStyle">Resolved shared stroke dash style.</param>
    /// <returns><c>true</c> when the preset represents a non-solid supported dash style.</returns>
    public static bool TryMapOfficePresetDash(string? presetDash, out OfficeStrokeDashStyle dashStyle) {
        dashStyle = OfficeStrokeDashStyle.Solid;
        string value = NormalizePresetDash(presetDash);
        switch (value) {
            case "dash":
            case "largedash":
            case "lgdash":
            case "systemdash":
            case "sysdash":
                dashStyle = OfficeStrokeDashStyle.Dash;
                return true;
            case "dot":
            case "systemdot":
            case "sysdot":
                dashStyle = OfficeStrokeDashStyle.Dot;
                return true;
            case "dashdot":
            case "largedashdot":
            case "lgdashdot":
            case "systemdashdot":
            case "sysdashdot":
                dashStyle = OfficeStrokeDashStyle.DashDot;
                return true;
            case "largedashdotdot":
            case "lgdashdotdot":
            case "systemdashdotdot":
            case "sysdashdotdot":
                dashStyle = OfficeStrokeDashStyle.DashDotDot;
                return true;
            default:
                return false;
        }
    }

    private static string NormalizePresetDash(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        return value!
            .Replace("-", string.Empty)
            .Replace("_", string.Empty)
            .Replace(" ", string.Empty)
            .Trim()
            .ToLowerInvariant();
    }
}
