using System;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>
/// Deterministic asset filename helpers for document read adapters.
/// </summary>
public static class OfficeDocumentAssetNaming {
    /// <summary>
    /// Builds a filesystem-safe filename from a stable asset id and optional extension.
    /// </summary>
    /// <param name="assetId">Stable asset identifier.</param>
    /// <param name="extension">Suggested file extension with or without a leading dot.</param>
    public static string BuildFileName(string assetId, string? extension) {
        if (assetId == null) throw new ArgumentNullException(nameof(assetId));

        string stem = SanitizeStem(assetId);
        string normalizedExtension = NormalizeExtension(extension);
        return normalizedExtension.Length == 0 ? stem : stem + normalizedExtension;
    }

    private static string SanitizeStem(string value) {
        var builder = new StringBuilder(value.Length);
        bool pendingSeparator = false;
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if ((c >= 'a' && c <= 'z') ||
                (c >= 'A' && c <= 'Z') ||
                (c >= '0' && c <= '9') ||
                c == '-' ||
                c == '_') {
                if (pendingSeparator && builder.Length > 0 && builder[builder.Length - 1] != '-' && builder[builder.Length - 1] != '_') {
                    builder.Append('-');
                }
                pendingSeparator = false;
                builder.Append(char.ToLowerInvariant(c));
            } else {
                pendingSeparator = true;
            }
        }

        while (builder.Length > 0 && builder[builder.Length - 1] == '-') {
            builder.Length--;
        }

        return builder.Length == 0 ? "asset" : builder.ToString();
    }

    private static string NormalizeExtension(string? extension) {
        if (string.IsNullOrWhiteSpace(extension)) {
            return string.Empty;
        }

        string value = extension!.Trim();
        if (value.StartsWith(".", StringComparison.Ordinal)) {
            value = value.Substring(1);
        }

        var builder = new StringBuilder(value.Length + 1);
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if ((c >= 'a' && c <= 'z') ||
                (c >= 'A' && c <= 'Z') ||
                (c >= '0' && c <= '9')) {
                builder.Append(char.ToLowerInvariant(c));
            }
        }

        return builder.Length == 0 ? string.Empty : "." + builder;
    }
}
