namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a table cell with optional Word-like column and row spanning.
/// </summary>
public sealed class PdfTableCell {
    /// <summary>Creates a table cell with text content, optional column/row spans, optional link metadata, images, and form fields.</summary>
    public PdfTableCell(string? text, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields = null, System.Collections.Generic.IEnumerable<PdfTableCellImage>? images = null, string? linkDestinationName = null, string? namedDestinationName = null) {
        Validate(columnSpan, rowSpan, linkUri, linkDestinationName, linkContents, namedDestinationName);
        Text = text ?? string.Empty;
        Runs = System.Array.AsReadOnly(new[] { TextRun.Normal(Text) });
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        LinkUri = linkUri;
        LinkDestinationName = linkDestinationName;
        NamedDestinationName = namedDestinationName;
        LinkContents = HasLinkTarget(linkUri, linkDestinationName) ? linkContents ?? Text : null;
        CheckBoxes = SnapshotCheckBoxes(checkBoxes, nameof(checkBoxes));
        FormFields = SnapshotFormFields(formFields, nameof(formFields));
        Images = SnapshotImages(images, nameof(images));
    }

    /// <summary>Creates a table cell with rich text runs, optional column/row spans, optional link metadata, images, and form fields.</summary>
    public PdfTableCell(System.Collections.Generic.IEnumerable<TextRun> runs, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields = null, System.Collections.Generic.IEnumerable<PdfTableCellImage>? images = null, string? linkDestinationName = null, string? namedDestinationName = null) {
        Guard.NotNull(runs, nameof(runs));
        Validate(columnSpan, rowSpan, linkUri, linkDestinationName, linkContents, namedDestinationName);
        var snapshot = new System.Collections.Generic.List<TextRun>();
        var text = new System.Text.StringBuilder();
        foreach (TextRun run in runs) {
            if (run is null) {
                throw new System.ArgumentException("Table cell text runs cannot contain null entries.", nameof(runs));
            }

            snapshot.Add(run);
            text.Append(run.Text);
        }

        Text = text.ToString();
        Runs = snapshot.AsReadOnly();
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        LinkUri = linkUri;
        LinkDestinationName = linkDestinationName;
        NamedDestinationName = namedDestinationName;
        LinkContents = HasLinkTarget(linkUri, linkDestinationName) ? linkContents ?? Text : null;
        CheckBoxes = SnapshotCheckBoxes(checkBoxes, nameof(checkBoxes));
        FormFields = SnapshotFormFields(formFields, nameof(formFields));
        Images = SnapshotImages(images, nameof(images));
    }

    /// <summary>Cell text content.</summary>
    public string Text { get; }

    /// <summary>Rich text runs for the cell. Plain text cells expose a single unstyled run.</summary>
    public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }

    /// <summary>Number of logical columns covered by this cell.</summary>
    public int ColumnSpan { get; }

    /// <summary>Number of logical rows covered by this cell.</summary>
    public int RowSpan { get; }

    /// <summary>Optional absolute URI or catalog-base-relative URI linked from this cell.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional PDF named destination linked from this cell.</summary>
    public string? LinkDestinationName { get; }

    /// <summary>Optional PDF named destination defined at this cell.</summary>
    public string? NamedDestinationName { get; }

    /// <summary>Optional PDF annotation contents metadata for the cell link.</summary>
    public string? LinkContents { get; }

    /// <summary>Simple AcroForm check boxes rendered inside this cell.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfTableCellCheckBox> CheckBoxes { get; }

    /// <summary>Simple AcroForm text and choice fields rendered inside this cell.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfTableCellFormField> FormFields { get; }

    /// <summary>Images rendered inside this cell.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfTableCellImage> Images { get; }

    /// <summary>Creates a single-column text cell.</summary>
    public static PdfTableCell TextCell(string? text, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null, string? namedDestinationName = null) => new PdfTableCell(text, linkUri: linkUri, linkContents: linkContents, linkDestinationName: linkDestinationName, namedDestinationName: namedDestinationName);

    /// <summary>Creates a single-column rich text cell.</summary>
    public static PdfTableCell RichTextCell(System.Collections.Generic.IEnumerable<TextRun> runs, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null, string? namedDestinationName = null) => new PdfTableCell(runs, linkUri: linkUri, linkContents: linkContents, linkDestinationName: linkDestinationName, namedDestinationName: namedDestinationName);

    /// <summary>Creates a cell spanning multiple logical columns.</summary>
    public static PdfTableCell Span(string? text, int columnSpan, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, linkDestinationName: linkDestinationName);

    /// <summary>Creates a rich text cell spanning multiple logical columns.</summary>
    public static PdfTableCell Span(System.Collections.Generic.IEnumerable<TextRun> runs, int columnSpan, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) => new PdfTableCell(runs, columnSpan, linkUri, linkContents, linkDestinationName: linkDestinationName);

    /// <summary>Creates a merged cell spanning logical columns and rows.</summary>
    public static PdfTableCell Merge(string? text, int columnSpan = 1, int rowSpan = 1, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, rowSpan, linkDestinationName: linkDestinationName);

    /// <summary>Creates a rich text merged cell spanning logical columns and rows.</summary>
    public static PdfTableCell Merge(System.Collections.Generic.IEnumerable<TextRun> runs, int columnSpan = 1, int rowSpan = 1, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) => new PdfTableCell(runs, columnSpan, linkUri, linkContents, rowSpan, linkDestinationName: linkDestinationName);

    /// <summary>Creates a table cell with rich text and simple AcroForm check boxes.</summary>
    public static PdfTableCell WithCheckBoxes(System.Collections.Generic.IEnumerable<TextRun> runs, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox> checkBoxes, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, string? linkDestinationName = null) => new PdfTableCell(runs, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, linkDestinationName: linkDestinationName);

    /// <summary>Creates a table cell with plain text and simple AcroForm check boxes.</summary>
    public static PdfTableCell WithCheckBoxes(string? text, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox> checkBoxes, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, string? linkDestinationName = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, linkDestinationName: linkDestinationName);

    /// <summary>Creates a table cell with rich text and simple AcroForm text or choice fields.</summary>
    public static PdfTableCell WithFormFields(System.Collections.Generic.IEnumerable<TextRun> runs, System.Collections.Generic.IEnumerable<PdfTableCellFormField> formFields, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, string? linkDestinationName = null) => new PdfTableCell(runs, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, formFields, linkDestinationName: linkDestinationName);

    /// <summary>Creates a table cell with plain text and simple AcroForm text or choice fields.</summary>
    public static PdfTableCell WithFormFields(string? text, System.Collections.Generic.IEnumerable<PdfTableCellFormField> formFields, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, string? linkDestinationName = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, formFields, linkDestinationName: linkDestinationName);

    /// <summary>Creates a table cell with rich text and images.</summary>
    public static PdfTableCell WithImages(System.Collections.Generic.IEnumerable<TextRun> runs, System.Collections.Generic.IEnumerable<PdfTableCellImage> images, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields = null, string? linkDestinationName = null) => new PdfTableCell(runs, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, formFields, images, linkDestinationName);

    /// <summary>Creates a table cell with plain text and images.</summary>
    public static PdfTableCell WithImages(string? text, System.Collections.Generic.IEnumerable<PdfTableCellImage> images, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields = null, string? linkDestinationName = null) => new PdfTableCell(text, columnSpan, linkUri, linkContents, rowSpan, checkBoxes, formFields, images, linkDestinationName);

    /// <summary>Returns a copy of this cell with a PDF named destination defined at the cell.</summary>
    public PdfTableCell WithNamedDestination(string? namedDestinationName) => new PdfTableCell(Runs, ColumnSpan, LinkUri, LinkContents, RowSpan, CheckBoxes, FormFields, Images, LinkDestinationName, namedDestinationName);

    internal PdfTableCell Clone() => new PdfTableCell(Runs, ColumnSpan, LinkUri, LinkContents, RowSpan, CheckBoxes, FormFields, Images, LinkDestinationName, NamedDestinationName);

    private static void Validate(int columnSpan, int rowSpan, string? linkUri, string? linkDestinationName, string? linkContents, string? namedDestinationName) {
        if (columnSpan < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(columnSpan), "Table cell column span must be at least 1.");
        }

        if (rowSpan < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(rowSpan), "Table cell row span must be at least 1.");
        }

        Guard.OptionalUriAction(linkUri, nameof(linkUri));

        if (linkUri != null && linkDestinationName != null) {
            throw new System.ArgumentException("A table cell link can target either a URI or a bookmark, not both.", nameof(linkDestinationName));
        }

        if (linkDestinationName != null) {
            Guard.NotNullOrWhiteSpace(linkDestinationName, nameof(linkDestinationName));
        }

        if (namedDestinationName != null) {
            Guard.NotNullOrWhiteSpace(namedDestinationName, nameof(namedDestinationName));
        }

        if (linkContents != null && !HasLinkTarget(linkUri, linkDestinationName)) {
            throw new System.ArgumentException("Link annotation contents require a link target.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }
    }

    private static bool HasLinkTarget(string? linkUri, string? linkDestinationName) => linkUri != null || linkDestinationName != null;

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfTableCellCheckBox> SnapshotCheckBoxes(System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes, string paramName) {
        if (checkBoxes == null) {
            return System.Array.AsReadOnly(System.Array.Empty<PdfTableCellCheckBox>());
        }

        var snapshot = new System.Collections.Generic.List<PdfTableCellCheckBox>();
        foreach (PdfTableCellCheckBox checkBox in checkBoxes) {
            if (checkBox == null) {
                throw new System.ArgumentException("Table cell check boxes cannot contain null entries.", paramName);
            }

            snapshot.Add(checkBox.Clone());
        }

        return snapshot.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfTableCellFormField> SnapshotFormFields(System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields, string paramName) {
        if (formFields == null) {
            return System.Array.AsReadOnly(System.Array.Empty<PdfTableCellFormField>());
        }

        var snapshot = new System.Collections.Generic.List<PdfTableCellFormField>();
        foreach (PdfTableCellFormField formField in formFields) {
            if (formField == null) {
                throw new System.ArgumentException("Table cell form fields cannot contain null entries.", paramName);
            }

            snapshot.Add(formField.Clone());
        }

        return snapshot.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfTableCellImage> SnapshotImages(System.Collections.Generic.IEnumerable<PdfTableCellImage>? images, string paramName) {
        if (images == null) {
            return System.Array.AsReadOnly(System.Array.Empty<PdfTableCellImage>());
        }

        var snapshot = new System.Collections.Generic.List<PdfTableCellImage>();
        foreach (PdfTableCellImage image in images) {
            if (image == null) {
                throw new System.ArgumentException("Table cell images cannot contain null entries.", paramName);
            }

            snapshot.Add(image.Clone());
        }

        return snapshot.AsReadOnly();
    }
}
