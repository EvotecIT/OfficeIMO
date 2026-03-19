using System.Text;
using System.Text.Json;

namespace OfficeIMO.MarkdownRenderer.Wpf;

internal static class MarkdownViewSupport {
    internal static OfficeIMO.MarkdownRenderer.MarkdownRendererOptions CreateEffectiveOptions(
        MarkdownViewPreset preset,
        string? baseHref,
        string? shellCss,
        Action<OfficeIMO.MarkdownRenderer.MarkdownRendererOptions>? configureRendererOptions) {
        var options = preset switch {
            MarkdownViewPreset.Relaxed => OfficeIMO.MarkdownRenderer.MarkdownRendererPresets.CreateRelaxed(),
            MarkdownViewPreset.StrictMinimal => OfficeIMO.MarkdownRenderer.MarkdownRendererPresets.CreateStrictMinimal(),
            _ => OfficeIMO.MarkdownRenderer.MarkdownRendererPresets.CreateStrict()
        };

        if (TryNormalizeBaseHref(baseHref, out var normalizedBaseHref)) {
            options.BaseHref = normalizedBaseHref;
        }

        if (!string.IsNullOrWhiteSpace(shellCss)) {
            options.ShellCss = AppendCss(options.ShellCss, shellCss);
        }

        configureRendererOptions?.Invoke(options);
        return options;
    }

    internal static string AppendCss(string? existing, string? additional) {
        var existingTrimmed = string.IsNullOrWhiteSpace(existing) ? string.Empty : (existing ?? string.Empty).Trim();
        var additionalTrimmed = string.IsNullOrWhiteSpace(additional) ? string.Empty : (additional ?? string.Empty).Trim();

        if (existingTrimmed.Length == 0) {
            return additionalTrimmed;
        }

        if (additionalTrimmed.Length == 0) {
            return existingTrimmed;
        }

        return new StringBuilder(existingTrimmed.Length + additionalTrimmed.Length + Environment.NewLine.Length)
            .Append(existingTrimmed)
            .Append(Environment.NewLine)
            .Append(additionalTrimmed)
            .ToString();
    }

    internal static bool TryNormalizeBaseHref(string? rawBaseHref, out string normalizedBaseHref) {
        normalizedBaseHref = string.Empty;
        if (string.IsNullOrWhiteSpace(rawBaseHref)) {
            return false;
        }

        if (!Uri.TryCreate((rawBaseHref ?? string.Empty).Trim(), UriKind.Absolute, out var parsed) || parsed == null) {
            return false;
        }

        normalizedBaseHref = parsed.AbsoluteUri;
        return true;
    }

    internal static bool TryGetClipboardText(string? webMessageAsJson, out string clipboardText) {
        clipboardText = string.Empty;
        if (string.IsNullOrWhiteSpace(webMessageAsJson)) {
            return false;
        }

        try {
            using var document = JsonDocument.Parse(webMessageAsJson ?? string.Empty);
            if (document.RootElement.ValueKind != JsonValueKind.Object) {
                return false;
            }

            if (!document.RootElement.TryGetProperty("type", out var typeElement)
                || !string.Equals(typeElement.GetString(), "omd.copy", StringComparison.Ordinal)) {
                return false;
            }

            if (!document.RootElement.TryGetProperty("text", out var textElement)
                || textElement.ValueKind != JsonValueKind.String) {
                return false;
            }

            var text = textElement.GetString();
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            clipboardText = text ?? string.Empty;
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    internal static bool TryGetExternalNavigationUri(string? rawUri, out Uri navigationUri) {
        navigationUri = null!;
        if (string.IsNullOrWhiteSpace(rawUri)
            || !Uri.TryCreate(rawUri, UriKind.Absolute, out var parsed)
            || parsed == null) {
            return false;
        }

        if (string.Equals(parsed.Scheme, "about", StringComparison.OrdinalIgnoreCase)
            || string.Equals(parsed.Scheme, "data", StringComparison.OrdinalIgnoreCase)
            || string.Equals(parsed.Scheme, "javascript", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        navigationUri = parsed;
        return true;
    }
}
