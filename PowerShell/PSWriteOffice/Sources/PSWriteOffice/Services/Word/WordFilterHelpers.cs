using System.Collections.Generic;
using System.Management.Automation;

namespace PSWriteOffice.Services.Word;

internal static class WordFilterHelpers
{
    internal static List<WildcardPattern> BuildPatterns(string[]? patterns)
    {
        var list = new List<WildcardPattern>();
        if (patterns == null || patterns.Length == 0)
        {
            return list;
        }

        foreach (var pattern in patterns)
        {
            if (!string.IsNullOrWhiteSpace(pattern))
            {
                list.Add(new WildcardPattern(pattern, WildcardOptions.IgnoreCase));
            }
        }

        return list;
    }

    internal static bool Matches(string? value, List<WildcardPattern> patterns)
    {
        if (patterns.Count == 0)
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return patterns.Exists(pattern => pattern.IsMatch(value));
    }
}
