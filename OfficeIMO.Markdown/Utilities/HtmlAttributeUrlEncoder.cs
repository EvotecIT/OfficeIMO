namespace OfficeIMO.Markdown;

internal static class HtmlAttributeUrlEncoder {
    internal static string Encode(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return System.Net.WebUtility.HtmlEncode(value.Replace(" ", "%20"));
    }

    internal static string EncodeSrcSet(string? srcSet) {
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return string.Empty;
        }

        var encodedCandidates = new System.Collections.Generic.List<string>();
        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(srcSet)) {
            string encodedUrl = Encode(candidate.Url);
            string encodedDescriptors = System.Net.WebUtility.HtmlEncode(candidate.Descriptor);
            encodedCandidates.Add(encodedDescriptors.Length == 0 ? encodedUrl : encodedUrl + " " + encodedDescriptors);
        }

        return string.Join(", ", encodedCandidates);
    }
}
