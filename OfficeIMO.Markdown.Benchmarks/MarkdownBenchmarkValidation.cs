using Markdig;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class MarkdownBenchmarkValidation {
    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();

    internal static void AssertCommonMarkEquivalent(
        string corpusName,
        string markdown,
        MarkdownReaderOptions officeOptions) {
        string officeHtml = NormalizeHtml(MarkdownReader.Parse(markdown, officeOptions).ToHtmlFragment());
        string markdigHtml = NormalizeHtml(Markdig.Markdown.ToHtml(markdown, MarkdigCommonMarkPipeline));
        if (string.Equals(officeHtml, markdigHtml, StringComparison.Ordinal)) {
            return;
        }

        int difference = FindFirstDifference(officeHtml, markdigHtml);
        throw new InvalidOperationException(
            $"CommonMark output differs for corpus '{corpusName}' at character {difference}. " +
            $"OfficeIMO length: {officeHtml.Length}; Markdig length: {markdigHtml.Length}. " +
            $"OfficeIMO: '{GetDifferenceWindow(officeHtml, difference)}'; " +
            $"Markdig: '{GetDifferenceWindow(markdigHtml, difference)}'.");
    }

    private static string NormalizeHtml(string html) {
        string normalized = html.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim();
        const string articlePrefix = "<article class=\"markdown-body\">";
        if (normalized.StartsWith(articlePrefix, StringComparison.Ordinal) &&
            normalized.EndsWith("</article>", StringComparison.Ordinal)) {
            normalized = normalized.Substring(
                articlePrefix.Length,
                normalized.Length - articlePrefix.Length - "</article>".Length);
        }

        normalized = Regex.Replace(normalized, "(<h[1-6]) id=\"[^\"]*\"", "$1");
        normalized = Regex.Replace(normalized, ">\\n+<", "><");
        normalized = normalized.Replace("<br />", "<br>", StringComparison.Ordinal).Trim();
        return NormalizeCollapsibleWhitespace(normalized);
    }

    private static string NormalizeCollapsibleWhitespace(string html) {
        string[] segments = Regex.Split(html, "(<pre(?:\\s[^>]*)?>.*?</pre>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        for (int index = 0; index < segments.Length; index += 2) {
            segments[index] = Regex.Replace(segments[index], "\\s+", " ");
            segments[index] = Regex.Replace(segments[index], "(<(?:p|li|h[1-6]|td|th|blockquote)(?:\\s[^>]*)?>) ", "$1", RegexOptions.IgnoreCase);
            segments[index] = Regex.Replace(segments[index], " </(p|li|h[1-6]|td|th|blockquote)>", "</$1>", RegexOptions.IgnoreCase);
            segments[index] = Regex.Replace(segments[index], " (?=<(?:article|blockquote|div|h[1-6]|ol|p|table|ul)(?:\\s|>))", string.Empty, RegexOptions.IgnoreCase);
        }

        return string.Concat(segments).Trim();
    }

    private static int FindFirstDifference(string left, string right) {
        int sharedLength = Math.Min(left.Length, right.Length);
        for (int index = 0; index < sharedLength; index++) {
            if (left[index] != right[index]) {
                return index;
            }
        }

        return sharedLength;
    }

    private static string GetDifferenceWindow(string value, int difference) {
        int start = Math.Max(0, difference - 30);
        int length = Math.Min(120, value.Length - start);
        return value.Substring(start, length).Replace("\n", "\\n", StringComparison.Ordinal);
    }
}
