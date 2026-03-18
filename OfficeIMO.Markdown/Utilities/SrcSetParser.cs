namespace OfficeIMO.Markdown;

internal readonly struct SrcSetCandidate {
    internal string Url { get; }
    internal string Descriptor { get; }

    internal SrcSetCandidate(string url, string descriptor) {
        Url = url ?? string.Empty;
        Descriptor = descriptor ?? string.Empty;
    }
}

internal static class SrcSetParser {
    internal static IReadOnlyList<SrcSetCandidate> Parse(string? srcSet) {
        var candidates = new List<SrcSetCandidate>();
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return candidates;
        }

        string value = srcSet!;
        int index = 0;
        while (index < value.Length) {
            SkipWhitespaceAndCommas(value, ref index);
            if (index >= value.Length) {
                break;
            }

            int urlStart = index;
            while (index < value.Length && !char.IsWhiteSpace(value[index])) {
                index++;
            }

            string url = value.Substring(urlStart, index - urlStart);
            int trailingCommaCount = 0;
            while (url.Length > 0 && url[url.Length - 1] == ',') {
                trailingCommaCount++;
                url = url.Substring(0, url.Length - 1);
            }

            url = url.Trim();
            if (url.Length == 0) {
                continue;
            }

            if (trailingCommaCount > 0) {
                candidates.Add(new SrcSetCandidate(url, string.Empty));
                continue;
            }

            SkipWhitespace(value, ref index);

            int descriptorStart = index;
            while (index < value.Length && value[index] != ',') {
                index++;
            }

            string descriptor = value.Substring(descriptorStart, index - descriptorStart).Trim();
            if (index < value.Length && value[index] == ',') {
                index++;
            }

            candidates.Add(new SrcSetCandidate(url, descriptor));
        }

        return candidates;
    }

    private static void SkipWhitespaceAndCommas(string value, ref int index) {
        while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) {
            index++;
        }
    }

    private static void SkipWhitespace(string value, ref int index) {
        while (index < value.Length && char.IsWhiteSpace(value[index])) {
            index++;
        }
    }
}
