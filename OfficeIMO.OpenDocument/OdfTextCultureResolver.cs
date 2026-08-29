using System.Globalization;

namespace OfficeIMO.OpenDocument;

/// <summary>Resolves document language metadata for deterministic text-case materialization.</summary>
internal static class OdfTextCultureResolver {
    internal static CultureInfo Resolve(string? language) {
        if (!string.IsNullOrWhiteSpace(language)) {
            try {
                return CultureInfo.GetCultureInfo(language!);
            } catch (CultureNotFoundException) {
                // Invalid document metadata falls back deterministically.
            }
        }

        return CultureInfo.InvariantCulture;
    }
}
