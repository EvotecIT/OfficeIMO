using System;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word;

/// <summary>
/// Provides utilities for adding inline formatted runs.
/// </summary>
public static class InlineRunHelper {
    private static readonly Regex _inlineRegex = new("(\\*\\*[^\\*]+\\*\\*|\\*[^\\*]+\\*|[^\\*]+)", RegexOptions.Singleline);
    private static readonly Regex _urlRegex = new("((?:https?|ftp)://[^\\s]+)", RegexOptions.IgnoreCase);

    /// <summary>
    /// Adds text runs to <paramref name="paragraph"/> parsing Markdown style bold and italic markers.
    /// </summary>
    /// <param name="paragraph">Paragraph to add runs to.</param>
    /// <param name="text">Input text containing optional <c>**bold**</c> or <c>*italic*</c> markers.</param>
    /// <param name="fontFamily">Optional font family for the runs.</param>
    public static void AddInlineRuns(WordParagraph paragraph, string text, string? fontFamily = null) {
        foreach (Match match in _inlineRegex.Matches(text)) {
            string token = match.Value;
            bool bold = token.StartsWith("**", StringComparison.Ordinal) && token.EndsWith("**");
            bool italic = !bold && token.StartsWith("*", StringComparison.Ordinal) && token.EndsWith("*");
            string value = bold ? token.Substring(2, token.Length - 4) :
                           italic ? token.Substring(1, token.Length - 2) : token;

            if (!bold && !italic && _urlRegex.IsMatch(value)) {
                int lastIndex = 0;
                foreach (Match urlMatch in _urlRegex.Matches(value)) {
                    if (urlMatch.Index > lastIndex) {
                        var textPart = value.Substring(lastIndex, urlMatch.Index - lastIndex);
                        var textRun = paragraph.AddFormattedText(textPart);
                        if (!string.IsNullOrEmpty(fontFamily)) {
                            textRun.SetFontFamily(fontFamily);
                        }
                    }

                    string url = urlMatch.Value;
                    var linkRun = paragraph.AddHyperLink(url, new Uri(url));
                    if (!string.IsNullOrEmpty(fontFamily)) {
                        linkRun.SetFontFamily(fontFamily);
                    }

                    lastIndex = urlMatch.Index + urlMatch.Length;
                }

                if (lastIndex < value.Length) {
                    var tailRun = paragraph.AddFormattedText(value.Substring(lastIndex));
                    if (!string.IsNullOrEmpty(fontFamily)) {
                        tailRun.SetFontFamily(fontFamily);
                    }
                }
            } else {
                var run = paragraph.AddFormattedText(value, bold, italic);
                if (!string.IsNullOrEmpty(fontFamily)) {
                    run.SetFontFamily(fontFamily);
                }
            }
        }
    }
}
