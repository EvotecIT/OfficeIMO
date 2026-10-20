using System;

namespace OfficeIMO.Drawing;

/// <summary>Shares the bounded MathML operator attribute grammar between import, layout and adapter diagnostics.</summary>
/// <remarks>Negative lengths, units other than em/px/percent, and CSS expressions remain unsupported.</remarks>
internal static class OfficeMathOperatorSpacing {
    /// <summary>Returns a recognized form using the MathML ASCII keyword spelling, or null for the inference default.</summary>
    internal static string? NormalizeForm(string? value) {
        string? form = value?.ToLowerInvariant();
        return form is "prefix" or "infix" or "postfix" ? form : null;
    }

    /// <summary>Retains a finite, nonnegative supported length without changing its authored unit.</summary>
    internal static string? NormalizeLength(string? value) =>
        TryParseLength(value, out _, out _) ? value!.Trim().ToLowerInvariant() : null;

    /// <summary>Resolves em against the core token size, percent against one dictionary side, and pixels against the host's scene units.</summary>
    internal static double ResolveLength(string? value, double dictionaryValue, double em, double pixelScale) {
        if (!TryParseLength(value, out double number, out string unit)) return dictionaryValue;
        double resolved = unit switch {
            "em" => number * em,
            "%" => (number / 100D) * dictionaryValue,
            "px" => number * pixelScale,
            _ => 0D
        };
        if (double.IsNaN(resolved) || double.IsInfinity(resolved))
            throw new ArgumentOutOfRangeException(nameof(value), "MathML operator spacing must fit inside finite drawing bounds.");
        return resolved;
    }

    private static bool TryParseLength(string? value, out double number, out string unit) {
        number = 0D;
        unit = string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string text = value!.Trim().ToLowerInvariant();
        int suffix = text.EndsWith("em", StringComparison.Ordinal) || text.EndsWith("px", StringComparison.Ordinal) ? 2
            : text.EndsWith("%", StringComparison.Ordinal) ? 1 : 0;
        unit = suffix == 0 ? string.Empty : text.Substring(text.Length - suffix);
        string numeric = suffix == 0 ? text : text.Substring(0, text.Length - suffix);
        return OfficeCssNumber.TryParse(numeric, out number) && number >= 0D
            && (suffix > 0 || number == 0D);
    }

}
