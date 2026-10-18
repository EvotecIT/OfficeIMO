using System;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

internal static class PowerPointObjectSearch
{
    public static bool MatchesShape(PowerPointShapeInfo info, Func<string, bool> matches)
    {
        if (info == null)
        {
            throw new ArgumentNullException(nameof(info));
        }

        if (matches == null)
        {
            throw new ArgumentNullException(nameof(matches));
        }

        return !string.IsNullOrEmpty(info.Text) && matches(info.Text!)
            || info.Shape is PowerPointTable table && matches(GetTableText(table));
    }

    public static Func<string, bool> CreateTextMatcher(string text, bool caseSensitive)
    {
        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return value => value?.IndexOf(text, comparison) >= 0;
    }

    public static Func<string, bool> CreateRegexMatcher(string pattern, bool caseSensitive)
    {
        var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        var regex = new Regex(pattern, options);
        return value => value != null && regex.IsMatch(value);
    }

    private static string GetTableText(PowerPointTable table)
    {
        var text = new StringBuilder();
        for (var row = 0; row < table.Rows; row++)
        {
            for (var column = 0; column < table.Columns; column++)
            {
                if (text.Length > 0)
                {
                    text.AppendLine();
                }

                text.Append(table.GetCell(row, column).Text);
            }
        }

        return text.ToString();
    }
}
