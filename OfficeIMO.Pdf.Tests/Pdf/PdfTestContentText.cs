using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Locates encoded text in content streams independently of TJ kerning and run grouping.</summary>
internal static class PdfTestContentText {
    private static readonly Regex TextArrayAdjustment = new(@">\s*-?(?:\d+\.?\d*|\.\d+)\s*<", RegexOptions.CultureInvariant);

    /// <summary>Matches a BaseFont/FontName entry with or without an embedded-subset tag (ABCDEF+).</summary>
    internal static string BaseFontPattern(string name) => "/BaseFont /(?:[A-Z]{6}\\+)?" + Regex.Escape(name);

    /// <summary>Removes an embedded-subset tag (ABCDEF+) from a font name.</summary>
    internal static string? WithoutSubsetTag(string? name) =>
        name != null && name.Length > 7 && name[6] == '+' && name.Take(6).All(static c => c >= 'A' && c <= 'Z') ? name.Substring(7) : name;

    /// <summary>Joins hex strings that a text-show array separates with numeric adjustments.</summary>
    internal static string JoinTextArrays(string content) => TextArrayAdjustment.Replace(content, string.Empty);
}
