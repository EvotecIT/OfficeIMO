namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private static bool IsName(XElement element, string expectedLocalName) {
        return string.Equals(element.Name.LocalName, expectedLocalName, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetAttribute(XElement element, string attributeName) {
        var attr = element.Attributes().FirstOrDefault(a => string.Equals(a.Name.LocalName, attributeName, StringComparison.OrdinalIgnoreCase));
        return attr?.Value ?? string.Empty;
    }

    private static string? TryGetFirstElementValue(XElement container, string localName) {
        foreach (var element in container.Elements()) {
            if (!string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) continue;

            var normalized = NormalizeWhitespace(element.Value);
            return normalized.Length == 0 ? null : normalized;
        }

        return null;
    }

    private static bool ContainsSpaceSeparatedToken(string? value, string token) {
        var source = value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(source)) return false;
        var parts = source.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts) {
            if (string.Equals(part, token, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static string RemoveFragmentAndQuery(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var trimmed = value.Trim();
        var hash = trimmed.IndexOf('#');
        if (hash >= 0) {
            trimmed = trimmed.Substring(0, hash);
        }

        var question = trimmed.IndexOf('?');
        if (question >= 0) {
            trimmed = trimmed.Substring(0, question);
        }

        if (trimmed.Length == 0) return string.Empty;

        try {
            return Uri.UnescapeDataString(trimmed);
        } catch {
            return trimmed;
        }
    }

    private static string CollapsePathSegments(string path) {
        var segments = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var stack = new List<string>(segments.Length);

        foreach (var segment in segments) {
            if (segment == ".") continue;
            if (segment == "..") {
                if (stack.Count == 0) return string.Empty;
                stack.RemoveAt(stack.Count - 1);
                continue;
            }

            stack.Add(segment);
        }

        return string.Join("/", stack);
    }

    private static string NormalizePath(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        return path.Replace('\\', '/').Trim();
    }

    private static bool TryNormalizeArchiveEntryPath(string path, out string normalizedPath) {
        normalizedPath = NormalizePath(path);
        if (normalizedPath.Length == 0 || normalizedPath.StartsWith("/", StringComparison.Ordinal)) return false;
        if (normalizedPath.Length >= 2 && normalizedPath[1] == ':') return false;
        if (Uri.TryCreate(normalizedPath, UriKind.Absolute, out _)) return false;

        string[] segments = normalizedPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (segments.Any(static segment => segment == "..")) return false;

        normalizedPath = CollapsePathSegments(normalizedPath);
        return normalizedPath.Length > 0;
    }

    private static string? NullIfWhiteSpace(string? value) {
        return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
    }

    private static string NormalizeWhitespace(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var sb = new StringBuilder(value.Length);
        bool previousWasWhitespace = false;
        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWasWhitespace) {
                    sb.Append(' ');
                    previousWasWhitespace = true;
                }
            } else {
                sb.Append(ch);
                previousWasWhitespace = false;
            }
        }

        return sb.ToString().Trim();
    }

    private static EpubReadOptions Normalize(EpubReadOptions? options) {
        var source = options ?? new EpubReadOptions();
        var normalized = new EpubReadOptions {
            MaxPackageBytes = source.MaxPackageBytes,
            MaxArchiveEntries = source.MaxArchiveEntries,
            MaxTotalUncompressedBytes = source.MaxTotalUncompressedBytes,
            MaxPackageMetadataBytes = source.MaxPackageMetadataBytes,
            MaxMetadataItems = source.MaxMetadataItems,
            MaxNavigationItems = source.MaxNavigationItems,
            MaxNavigationDepth = source.MaxNavigationDepth,
            MaxChapters = source.MaxChapters,
            MaxChapterBytes = source.MaxChapterBytes,
            MaxTotalRawHtmlBytes = source.MaxTotalRawHtmlBytes,
            IncludeRawHtml = source.IncludeRawHtml,
            IncludeResourceData = source.IncludeResourceData,
            MaxResources = source.MaxResources,
            MaxResourceBytes = source.MaxResourceBytes,
            MaxTotalResourceBytes = source.MaxTotalResourceBytes,
            DeterministicOrder = source.DeterministicOrder,
            PreferSpineOrder = source.PreferSpineOrder,
            IncludeNonLinearSpineItems = source.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = source.FallbackToHtmlScan
        };

        if (normalized.MaxPackageBytes < 1) normalized.MaxPackageBytes = 1;
        if (normalized.MaxArchiveEntries < 1) normalized.MaxArchiveEntries = 1;
        if (normalized.MaxTotalUncompressedBytes < 1) normalized.MaxTotalUncompressedBytes = 1;
        if (normalized.MaxPackageMetadataBytes < 1) normalized.MaxPackageMetadataBytes = 1;
        if (normalized.MaxMetadataItems < 1) normalized.MaxMetadataItems = 1;
        if (normalized.MaxNavigationItems < 1) normalized.MaxNavigationItems = 1;
        if (normalized.MaxNavigationDepth < 1) normalized.MaxNavigationDepth = 1;
        if (normalized.MaxChapters < 1) normalized.MaxChapters = 1;
        if (normalized.MaxChapterBytes.HasValue && normalized.MaxChapterBytes.Value < 1) {
            normalized.MaxChapterBytes = 1;
        }
        if (normalized.MaxTotalRawHtmlBytes < 1) normalized.MaxTotalRawHtmlBytes = 1;
        if (normalized.MaxResources < 1) normalized.MaxResources = 1;
        if (normalized.MaxResourceBytes < 1) normalized.MaxResourceBytes = 1;
        if (normalized.MaxTotalResourceBytes < 1) normalized.MaxTotalResourceBytes = 1;

        return normalized;
    }

}
