using Markdig;
using OfficeIMO.Benchmarks;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class MarkdownBenchmarkValidation {
    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();

    internal static void AssertCommonMarkEquivalent(
        string corpusName,
        string markdown,
        MarkdownLoadOptions officeOptions,
        MarkdownHtmlExportOptions officeHtmlOptions) {
        string officeHtml = NormalizeHtml(MarkdownBenchmarkParsing.ParseSemantic(markdown, officeOptions).ToHtml(officeHtmlOptions));
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

    /// <summary>
    /// Checks the HTML the measured OfficeIMO and Markdig lanes produce for one corpus: after whitespace normalization
    /// both must render the same CommonMark HTML. Records the result for both libraries and throws on a mismatch.
    /// </summary>
    internal static void ValidateMeasuredLanes(string corpusName, string caseName, string officeHtml, string markdigHtml) {
        string office = NormalizeHtml(officeHtml);
        string markdig = NormalizeHtml(markdigHtml);
        string detail = $"CorpusName={corpusName}; normalized HTML characters OfficeIMO={office.Length}, Markdig={markdig.Length}";
        if (office.Length > 0 && string.Equals(office, markdig, StringComparison.Ordinal)) {
            BenchmarkOutputValidationLog.Passed("OfficeIMO", detail + "; identical CommonMark HTML", caseName);
            BenchmarkOutputValidationLog.Passed("Markdig", detail + "; identical CommonMark HTML", caseName);
            return;
        }

        int difference = FindFirstDifference(office, markdig);
        string message = $"CommonMark output of the measured lanes differs for corpus '{corpusName}' at character {difference}. " +
            $"OfficeIMO: '{GetDifferenceWindow(office, difference)}'; Markdig: '{GetDifferenceWindow(markdig, difference)}'.";
        BenchmarkOutputValidationLog.Failed("OfficeIMO", detail + "; " + message, caseName);
        BenchmarkOutputValidationLog.Failed("Markdig", detail + "; " + message, caseName);
        throw new InvalidOperationException(message);
    }

    internal static string NormalizeHtml(string html) {
        string normalized = html.Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
        normalized = Regex.Replace(normalized, ">\\n+<", "><");
        normalized = normalized.Replace("<br />", "<br>").Trim();
        return NormalizeCollapsibleWhitespace(normalized);
    }

    internal static MarkdownHtmlExportOptions CreateOfficeCommonMarkHtmlOptions() => new MarkdownHtmlExportOptions {
        AutoHeadingIdentifiers = false,
        BodyClass = null
    };

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
        return value.Substring(start, length).Replace("\n", "\\n");
    }
}
