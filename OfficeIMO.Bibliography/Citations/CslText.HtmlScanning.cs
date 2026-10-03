using System.Net;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    /// <summary>Scans owned escaped HTML without a host-dependent regex timeout.</summary>
    private static IEnumerable<(string Value, bool Tag)> HtmlParts(string html, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int copied = 0, position = 0;
        while (position < html.Length) {
            CslNumberSyntax.Check(position, token);
            if (html[position] != '<') { position++; continue; }
            int end = position + 1;
            while (end < html.Length && html[end] != '>') { CslNumberSyntax.Check(end, token); end++; }
            // No later opening angle can complete a tag if there is no closing angle.
            if (end == html.Length) break;
            yield return (html.Substring(copied, position - copied), false);
            yield return (html.Substring(position, end - position + 1), true);
            position = copied = end + 1;
        }
        yield return (html.Substring(copied), false);
    }

    internal static string HtmlPlain(string html, CancellationToken token = default) {
        var text = new StringBuilder(html.Length);
        foreach (var part in HtmlParts(html, token)) if (!part.Tag) text.Append(part.Value);
        return WebUtility.HtmlDecode(text.ToString());
    }

    private static string TransformHtml(string html, Func<string, bool, string> transform, CancellationToken token = default) {
        var output = new StringBuilder(html.Length);
        var tags = new Stack<bool>();
        bool protectedCase = false;
        foreach (var part in HtmlParts(html, token)) {
            token.ThrowIfCancellationRequested();
            string value = part.Value;
            if (value.StartsWith("</", StringComparison.Ordinal)) {
                if (tags.Count > 0) protectedCase = tags.Pop();
                output.Append(value);
            } else if (value.StartsWith("<", StringComparison.Ordinal)) {
                tags.Push(protectedCase);
                protectedCase |= value.Contains("class=\"nocase\"") || value.Contains("class=\"nodecor\"") ||
                    value.Contains("class=\"nocase nodecor\"") || value == "<sup>" || value == "<sub>" || value.Contains("font-variant:small-caps");
                output.Append(value);
            } else output.Append(Escape(transform(WebUtility.HtmlDecode(value), protectedCase)));
        }
        return output.ToString();
    }
}
