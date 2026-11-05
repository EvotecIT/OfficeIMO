using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Captures the PowerShell case policy for native redaction preview and application.</summary>
internal static class PdfRedactionCommandOptions {
    internal static PdfRedactionSearchOptions Create(string[]? text, string[]? patterns, bool caseSensitive) {
        var search = new PdfRedactionSearchOptions {
            MatchCase = caseSensitive,
            RegexOptions = RegexOptions.CultureInvariant | (caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase)
        };
        if (text != null) search.AddLiteral(text);
        if (patterns != null) search.AddRegex(patterns);
        return search;
    }
}
