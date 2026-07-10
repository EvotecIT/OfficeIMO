namespace OfficeIMO.OpenDocument;

/// <summary>ODT block kinds exposed in document reading order.</summary>
public enum OdtContentBlockKind {
    /// <summary>Normal paragraph.</summary>
    Paragraph,
    /// <summary>Outline heading.</summary>
    Heading,
    /// <summary>Table.</summary>
    Table
}

/// <summary>A paragraph, heading, or table in ODT reading order.</summary>
public sealed class OdtContentBlock {
    private OdtContentBlock(OdtContentBlockKind kind, OdtParagraph? paragraph, OdtTable? table) {
        Kind = kind;
        Paragraph = paragraph;
        Table = table;
    }

    /// <summary>Block kind.</summary>
    public OdtContentBlockKind Kind { get; }
    /// <summary>Paragraph or heading value when applicable.</summary>
    public OdtParagraph? Paragraph { get; }
    /// <summary>Table value when applicable.</summary>
    public OdtTable? Table { get; }

    internal static OdtContentBlock FromParagraph(OdtParagraph paragraph) => new OdtContentBlock(
        paragraph.IsHeading ? OdtContentBlockKind.Heading : OdtContentBlockKind.Paragraph, paragraph, null);

    internal static OdtContentBlock FromTable(OdtTable table) => new OdtContentBlock(OdtContentBlockKind.Table, null, table);
}
