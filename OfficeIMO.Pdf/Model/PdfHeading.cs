namespace OfficeIMO.Pdf;

/// <summary>Describes a heading added to the document flow and outline.</summary>
public sealed class PdfHeading {
    internal PdfHeading(string text, int level) { Text = text; Level = level; }

    /// <summary>Heading text, also used as its outline title.</summary>
    public string Text { get; }

    /// <summary>Heading level from 1 through 9.</summary>
    public int Level { get; }
}
