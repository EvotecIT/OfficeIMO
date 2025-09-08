using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Common header transform helpers for converting property names to human-friendly headers.
/// </summary>
public static class HeaderTransforms {
    private static readonly Regex SplitPascal = new Regex("(?<!^)(?=[A-Z])", RegexOptions.Compiled);
    private static readonly string[] Acronyms = new[] { "DNS", "SPF", "DKIM", "DMARC", "BIMI", "TLS", "SSL", "MX", "PTR", "WHOIS", "URL", "HTML", "PDF", "CPU", "GPU" };

    /// <summary>
    /// Converts PascalCase or snake_case to a spaced title and uppercases known acronyms.
    /// </summary>
    public static string Pretty(string name) {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        // snake_case to spaces
        name = name.Replace('_', ' ');
        // split PascalCase
        if (!name.Contains(' ')) name = SplitPascal.Replace(name, " ");
        // Trim
        name = name.Trim();
        // Uppercase known acronyms
        foreach (var ac in Acronyms) {
            name = Regex.Replace(name, $"\\b{ac.Substring(0,1)}{ac.Substring(1).ToLower()}\\b|\\b{ac.ToLower()}\\b|\\b{ac}\\b", ac, RegexOptions.IgnoreCase);
        }
        // Title case words that are not fully uppercase
        var parts = name.Split(' ');
        var sb = new StringBuilder();
        for (int i = 0; i < parts.Length; i++) {
            var w = parts[i];
            if (IsAllUpper(w)) { sb.Append(w); }
            else { sb.Append(char.ToUpperInvariant(w[0])); if (w.Length > 1) sb.Append(w.Substring(1).ToLowerInvariant()); }
            if (i < parts.Length - 1) sb.Append(' ');
        }
        return sb.ToString();
    }

    private static bool IsAllUpper(string s) {
        for (int i = 0; i < s.Length; i++) if (char.IsLetter(s[i]) && !char.IsUpper(s[i])) return false; return true;
    }
}

