namespace OfficeIMO.OneNote.Markdown;

/// <summary>Normalizes native OneNote text for safe text, Markdown, HTML, Reader, and PDF projection.</summary>
public static class OneNoteTextProjection {
    /// <summary>
    /// Converts OneNote/RichEdit layout controls into line breaks and replaces invalid Unicode
    /// controls, noncharacters, and unpaired surrogates. The typed source model is not modified.
    /// </summary>
    /// <param name="value">Native OneNote text.</param>
    /// <returns>Projection-safe Unicode text.</returns>
    public static string Normalize(string? value) => OneNoteTextNormalizer.Normalize(value);
}
