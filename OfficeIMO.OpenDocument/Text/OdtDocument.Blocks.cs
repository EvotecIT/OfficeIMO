namespace OfficeIMO.OpenDocument;

public sealed partial class OdtDocument {
    /// <summary>Paragraphs and headings in document body order.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => TextBody
        .Descendants()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Where(element => !element.Ancestors(OdfNamespaces.Text + "tracked-changes").Any())
        .Select(element => new OdtParagraph(this, element))
        .ToList();

    /// <summary>Top-level tables in the document body.</summary>
    public IReadOnlyList<OdtTable> Tables => TextBody.Elements(OdfNamespaces.Table + "table")
        .Select(element => new OdtTable(this, element)).ToList();

    /// <summary>Top-level lists in the document body.</summary>
    public IReadOnlyList<OdtList> Lists => TextBody.Elements(OdfNamespaces.Text + "list")
        .Select(element => new OdtList(this, element)).ToList();

    /// <summary>Top-level named sections in the document body.</summary>
    public IReadOnlyList<OdtSection> Sections => TextBody.Elements(OdfNamespaces.Text + "section")
        .Select(element => new OdtSection(this, element)).ToList();

    /// <summary>Paragraph, heading, and table blocks in reading order without duplicating table-cell paragraphs.</summary>
    public IReadOnlyList<OdtContentBlock> ContentBlocks => EnumerateContentBlocks(TextBody).ToList();

    /// <summary>Adds a paragraph to the document body.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        var element = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(element, text);
        TextBody.Add(element);
        MarkPartDirty("content.xml");
        return new OdtParagraph(this, element);
    }

    /// <summary>Adds a heading with an outline level from 1 through 10.</summary>
    public OdtParagraph AddHeading(string text, int level = 1) {
        if (level < 1 || level > 10) throw new ArgumentOutOfRangeException(nameof(level), "Heading level must be between 1 and 10.");
        var element = new XElement(OdfNamespaces.Text + "h",
            new XAttribute(OdfNamespaces.Text + "outline-level", level));
        OdfTextCodec.Append(element, text);
        TextBody.Add(element);
        MarkPartDirty("content.xml");
        return new OdtParagraph(this, element);
    }

    /// <summary>Adds an ordered or unordered list.</summary>
    public OdtList AddList(bool ordered = false) {
        string styleName = OdfListStyleStore.Create(this, ordered);
        var element = new XElement(OdfNamespaces.Text + "list",
            new XAttribute(OdfNamespaces.Text + "style-name", styleName));
        TextBody.Add(element);
        MarkPartDirty("content.xml");
        return new OdtList(this, element);
    }

    /// <summary>Adds a table initialized with the requested row and column count.</summary>
    public OdtTable AddTable(int rows, int columns, string? name = null) {
        if (rows < 1) throw new ArgumentOutOfRangeException(nameof(rows));
        if (columns < 1) throw new ArgumentOutOfRangeException(nameof(columns));
        string tableName = string.IsNullOrWhiteSpace(name) ? NextTableName() : name!;
        var table = new XElement(OdfNamespaces.Table + "table",
            new XAttribute(OdfNamespaces.Table + "name", tableName),
            new XElement(OdfNamespaces.Table + "table-column",
                new XAttribute(OdfNamespaces.Table + "number-columns-repeated", columns)));
        for (int row = 0; row < rows; row++) {
            var rowElement = new XElement(OdfNamespaces.Table + "table-row");
            for (int column = 0; column < columns; column++) {
                rowElement.Add(OdtTableCell.CreateElement());
            }
            table.Add(rowElement);
        }
        TextBody.Add(table);
        MarkPartDirty("content.xml");
        return new OdtTable(this, table);
    }

    /// <summary>Adds a named section to the document body.</summary>
    public OdtSection AddSection(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(name));
        var element = new XElement(OdfNamespaces.Text + "section", new XAttribute(OdfNamespaces.Text + "name", name));
        TextBody.Add(element);
        MarkPartDirty("content.xml");
        return new OdtSection(this, element);
    }

    /// <summary>Adds an empty paragraph that starts on a new page.</summary>
    public OdtParagraph AddPageBreak() {
        OdtParagraph paragraph = AddParagraph();
        paragraph.PageBreakBefore = true;
        return paragraph;
    }

    internal IEnumerable<XElement> BodyBlocks() {
        foreach (XElement element in TextBody.Descendants()) {
            if (element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h" || element.Name == OdfNamespaces.Table + "table") {
                if (element.Ancestors(OdfNamespaces.Text + "tracked-changes").Any()) continue;
                yield return element;
            }
        }
    }

    private IEnumerable<OdtContentBlock> EnumerateContentBlocks(XElement container, bool? orderedList = null, int listLevel = 0) {
        foreach (XElement element in container.Elements()) {
            if (element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h") {
                yield return OdtContentBlock.FromParagraph(new OdtParagraph(this, element), orderedList.HasValue, orderedList, listLevel);
            } else if (element.Name == OdfNamespaces.Table + "table") {
                yield return OdtContentBlock.FromTable(new OdtTable(this, element));
            } else if (element.Name == OdfNamespaces.Text + "list") {
                string? styleName = (string?)element.Attribute(OdfNamespaces.Text + "style-name");
                bool ordered = OdfListStyleStore.IsOrdered(this, styleName);
                foreach (OdtContentBlock block in EnumerateContentBlocks(element, ordered, orderedList.HasValue ? listLevel + 1 : listLevel)) yield return block;
            } else if (element.Name == OdfNamespaces.Text + "section" || element.Name == OdfNamespaces.Text + "list-item" ||
                       element.Name == OdfNamespaces.Text + "list-header") {
                foreach (OdtContentBlock block in EnumerateContentBlocks(element, orderedList, listLevel)) yield return block;
            }
        }
    }

    private string NextTableName() {
        var names = new HashSet<string>(TextBody.Descendants(OdfNamespaces.Table + "table")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Table + "name"))
            .Where(value => !string.IsNullOrEmpty(value))!, StringComparer.Ordinal);
        int index = 1;
        string name;
        do { name = "Table" + index++.ToString(CultureInfo.InvariantCulture); } while (names.Contains(name));
        return name;
    }
}
