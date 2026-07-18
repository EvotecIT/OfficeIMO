namespace OfficeIMO.Pdf;

/// <summary>
/// Canonical content-stream operator registry shared by diagnostics and debugger views.
/// Lexical framing remains owned by <see cref="PdfContentStreamInterpreter"/>.
/// </summary>
internal static class PdfContentOperators {
    private static readonly HashSet<string> StandardOperators = new HashSet<string>(StringComparer.Ordinal) {
        "q", "Q", "cm", "w", "J", "j", "d", "gs", "CS", "cs", "SC", "SCN", "sc", "scn", "G", "g", "RG", "rg", "K", "k",
        "m", "l", "c", "v", "y", "h", "re", "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n", "W", "W*", "sh",
        "BT", "ET", "Tc", "Tw", "Tz", "TL", "Tf", "Tr", "Ts", "Td", "TD", "Tm", "T*", "Tj", "TJ", "'", "\"",
        "Do", "BI", "ID", "EI", "BMC", "BDC", "EMC", "BX", "EX", "M", "ri", "i", "MP", "DP", "d0", "d1"
    };

    internal static bool IsStandard(string value) => StandardOperators.Contains(value);
}
