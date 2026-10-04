using System.Text.Json;

namespace OfficeIMO.Bibliography;

/// <summary>Compares structured name parts without interpreting literal names or full given names as initials.</summary>
internal static class CslNameIdentity {
    private static readonly string[] Parts = { "literal", "family", "given", "suffix", "dropping-particle", "non-dropping-particle" };

    internal static string Read(JsonElement name, bool includeCommaSuffix, CancellationToken token) {
        var result = new StringBuilder();
        foreach (string part in Parts) {
            token.ThrowIfCancellationRequested();
            string value = name.TryGetProperty(part, out JsonElement field) ? field.ToString() : string.Empty;
            if (part == "given") value = Initials(value, token);
            Append(result, value);
        }
        if (includeCommaSuffix) Append(result, name.TryGetProperty("comma-suffix", out JsonElement comma) ? comma.ToString() : string.Empty);
        return result.ToString();
    }

    // Length framing keeps a separator in one source name part from becoming another part.
    private static void Append(StringBuilder output, string value) =>
        output.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);

    private static string Initials(string source, CancellationToken token) {
        var output = new StringBuilder(source.Length);
        int index = 0;
        bool first = true;
        while (index < source.Length) {
            token.ThrowIfCancellationRequested();
            while (index < source.Length && char.IsWhiteSpace(source[index])) {
                if ((index & 255) == 0) token.ThrowIfCancellationRequested();
                index++;
            }
            if (index == source.Length) break;
            if (!char.IsLetter(source, index)) return source;
            int start = index;
            index += char.IsSurrogatePair(source, index) ? 2 : 1;
            int marks = 0;
            while (index < source.Length) {
                if ((marks++ & 255) == 0) token.ThrowIfCancellationRequested();
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(source, index);
                if (category != UnicodeCategory.NonSpacingMark && category != UnicodeCategory.SpacingCombiningMark && category != UnicodeCategory.EnclosingMark) break;
                index += char.IsSurrogatePair(source, index) ? 2 : 1;
            }
            // An adjacent letter makes this a full name, rather than a sequence of initials.
            if (index < source.Length && source[index] != '.' && source[index] != '-' && !char.IsWhiteSpace(source[index])) return source;
            if (!first && output[output.Length - 1] != '-') output.Append(' ');
            output.Append(source, start, index - start).Append('.');
            first = false;
            if (index < source.Length && source[index] == '.') index++;
            while (index < source.Length && char.IsWhiteSpace(source[index])) {
                if ((index & 255) == 0) token.ThrowIfCancellationRequested();
                index++;
            }
            if (index < source.Length && source[index] == '-') { output.Append('-'); index++; }
        }
        return output.Length == 0 || output[output.Length - 1] == '-' ? source : output.ToString();
    }
}
