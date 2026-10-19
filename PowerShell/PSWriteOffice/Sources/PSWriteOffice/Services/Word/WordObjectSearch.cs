using System;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal static class WordObjectSearch
{
    public static bool MatchesTable(WordTable table, Func<string, bool> matches)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        if (matches == null)
        {
            throw new ArgumentNullException(nameof(matches));
        }

        return table.Rows
            .SelectMany(row => row.Cells)
            .SelectMany(cell => cell.Paragraphs)
            .Any(paragraph => matches(paragraph.Text ?? string.Empty));
    }

    public static bool MatchesList(WordList list, Func<string, bool> matches)
    {
        if (list == null)
        {
            throw new ArgumentNullException(nameof(list));
        }
        if (matches == null)
        {
            throw new ArgumentNullException(nameof(matches));
        }

        return list.ListItems.Any(item => matches(item.Text ?? string.Empty));
    }

    public static Func<string, bool> CreateTextMatcher(string text, bool caseSensitive)
    {
        if (string.IsNullOrEmpty(text))
        {
            throw new ArgumentException("Provide text to find.", nameof(text));
        }

        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return value => value.IndexOf(text, comparison) >= 0;
    }

    public static Func<string, bool> CreateRegexMatcher(string pattern, bool caseSensitive)
    {
        if (string.IsNullOrEmpty(pattern))
        {
            throw new ArgumentException("Provide a regex pattern to find.", nameof(pattern));
        }

        var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        var regex = new Regex(pattern, options);
        return value => regex.IsMatch(value);
    }
}
