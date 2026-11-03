namespace OfficeIMO.Pdf;

/// <summary>Describes a bulleted or numbered list added to the document flow.</summary>
public sealed class PdfList {
    internal PdfList(int itemCount, bool isNumbered) { ItemCount = itemCount; IsNumbered = isNumbered; }

    /// <summary>Number of list items.</summary>
    public int ItemCount { get; }

    /// <summary>True for a numbered list, false for a bulleted list.</summary>
    public bool IsNumbered { get; }
}
