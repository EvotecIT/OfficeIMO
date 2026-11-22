namespace OfficeIMO.Pdf;

/// <summary>
/// The characters that end a line in laid-out text: a carriage return, a line feed or both.
/// Converters that inspect source lines share this rule with the writer.
/// </summary>
internal static class PdfTextLineEnds {
    internal static bool IsLineEnd(char character) => character == '\r' || character == '\n';

    // A non-allocating scan: string.Contains(char) is unavailable on netstandard2.0/net472, and
    // a LINQ scan allocates an enumerator per call on the layout hot path.
    internal static bool Contains(string text) {
        for (int i = 0; i < text.Length; i++) {
            if (IsLineEnd(text[i])) return true;
        }

        return false;
    }
}
