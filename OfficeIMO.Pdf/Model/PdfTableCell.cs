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
        Paragraphs = System.Array.AsReadOnly(System.Array.Empty<PdfTableCellParagraph>());
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
        Paragraphs = System.Array.AsReadOnly(System.Array.Empty<PdfTableCellParagraph>());
    }

    internal PdfTableCell(System.Collections.Generic.IEnumerable<TextRun> runs, System.Collections.Generic.IEnumerable<PdfTableCellParagraph>? paragraphs, int columnSpan = 1, string? linkUri = null, string? linkContents = null, int rowSpan = 1, System.Collections.Generic.IEnumerable<PdfTableCellCheckBox>? checkBoxes = null, System.Collections.Generic.IEnumerable<PdfTableCellFormField>? formFields = null, System.Collections.Generic.IEnumerable<PdfTableCellImage>? images = null, string? linkDestinationName = null, string? namedDestinationName = null, bool noWrap = false) {
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
        Paragraphs = SnapshotParagraphs(paragraphs, nameof(paragraphs));
        NoWrap = noWrap;
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

    internal System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> Paragraphs { get; }

    internal bool NoWrap { get; }

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
    public PdfTableCell WithNamedDestination(string? namedDestinationName) => new PdfTableCell(Runs, Paragraphs, ColumnSpan, LinkUri, LinkContents, RowSpan, CheckBoxes, FormFields, Images, LinkDestinationName, namedDestinationName, NoWrap);

    /// <summary>
    /// Returns a copy that keeps each cell paragraph on one visual line. When the containing
    /// table enables text shrinking, the renderer reduces the font before clipping.
    /// </summary>
    public PdfTableCell WithNoWrap(bool noWrap = true) => new PdfTableCell(Runs, Paragraphs, ColumnSpan, LinkUri, LinkContents, RowSpan, CheckBoxes, FormFields, Images, LinkDestinationName, NamedDestinationName, noWrap);

    internal PdfTableCell Clone() => new PdfTableCell(Runs, Paragraphs, ColumnSpan, LinkUri, LinkContents, RowSpan, CheckBoxes, FormFields, Images, LinkDestinationName, NamedDestinationName, NoWrap);

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

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfTableCellParagraph> SnapshotParagraphs(System.Collections.Generic.IEnumerable<PdfTableCellParagraph>? paragraphs, string paramName) {
        if (paragraphs == null) {
            return System.Array.AsReadOnly(System.Array.Empty<PdfTableCellParagraph>());
        }

        var snapshot = new System.Collections.Generic.List<PdfTableCellParagraph>();
        foreach (PdfTableCellParagraph paragraph in paragraphs) {
            if (paragraph == null) {
                throw new System.ArgumentException("Table cell paragraphs cannot contain null entries.", paramName);
            }

            snapshot.Add(paragraph.Clone());
        }

        return snapshot.AsReadOnly();
    }
}

internal sealed class PdfTableCellParagraph {
    public PdfTableCellParagraph(System.Collections.Generic.IEnumerable<TextRun> runs, double spacingAfter = 0D, PdfAlign? align = null, double spacingBefore = 0D, double leftIndent = 0D, double rightIndent = 0D, double firstLineIndent = 0D, double? lineHeight = null, double? defaultTabStopWidth = null, System.Collections.Generic.IEnumerable<PdfTabStop>? tabStops = null) {
        Guard.NotNull(runs, nameof(runs));
        if (spacingBefore < 0 || double.IsNaN(spacingBefore) || double.IsInfinity(spacingBefore)) {
            throw new System.ArgumentOutOfRangeException(nameof(spacingBefore), "Table cell paragraph spacing must be a non-negative finite value.");
        }

        if (spacingAfter < 0 || double.IsNaN(spacingAfter) || double.IsInfinity(spacingAfter)) {
            throw new System.ArgumentOutOfRangeException(nameof(spacingAfter), "Table cell paragraph spacing must be a non-negative finite value.");
        }

        if (leftIndent < 0 || double.IsNaN(leftIndent) || double.IsInfinity(leftIndent)) {
            throw new System.ArgumentOutOfRangeException(nameof(leftIndent), "Table cell paragraph left indent must be a non-negative finite value.");
        }

        if (rightIndent < 0 || double.IsNaN(rightIndent) || double.IsInfinity(rightIndent)) {
            throw new System.ArgumentOutOfRangeException(nameof(rightIndent), "Table cell paragraph right indent must be a non-negative finite value.");
        }

        if (double.IsNaN(firstLineIndent) || double.IsInfinity(firstLineIndent)) {
            throw new System.ArgumentOutOfRangeException(nameof(firstLineIndent), "Table cell paragraph first line indent must be a finite value.");
        }

        if (lineHeight.HasValue && (lineHeight.Value <= 0 || double.IsNaN(lineHeight.Value) || double.IsInfinity(lineHeight.Value))) {
            throw new System.ArgumentOutOfRangeException(nameof(lineHeight), "Table cell paragraph line height must be a positive finite value.");
        }

        if (defaultTabStopWidth.HasValue && (defaultTabStopWidth.Value <= 0 || double.IsNaN(defaultTabStopWidth.Value) || double.IsInfinity(defaultTabStopWidth.Value))) {
            throw new System.ArgumentOutOfRangeException(nameof(defaultTabStopWidth), "Table cell paragraph default tab stop width must be a positive finite value.");
        }

        var snapshot = new System.Collections.Generic.List<TextRun>();
        foreach (TextRun run in runs) {
            if (run is null) {
                throw new System.ArgumentException("Table cell paragraph runs cannot contain null entries.", nameof(runs));
            }

            snapshot.Add(run);
        }

        var tabStopSnapshot = new System.Collections.Generic.List<PdfTabStop>();
        if (tabStops != null) {
            foreach (PdfTabStop tabStop in tabStops) {
                if (tabStop is null) {
                    throw new System.ArgumentException("Table cell paragraph tab stops cannot contain null entries.", nameof(tabStops));
                }

                tabStopSnapshot.Add(tabStop.Clone());
            }
        }

        Runs = snapshot.AsReadOnly();
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        Align = align;
        LeftIndent = leftIndent;
        RightIndent = rightIndent;
        FirstLineIndent = firstLineIndent;
        LineHeight = lineHeight;
        DefaultTabStopWidth = defaultTabStopWidth;
        TabStops = tabStopSnapshot.AsReadOnly();
    }

    public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }

    public double SpacingBefore { get; }

    public double SpacingAfter { get; }

    public PdfAlign? Align { get; }

    public double LeftIndent { get; }

    public double RightIndent { get; }

    public double FirstLineIndent { get; }

    public double? LineHeight { get; }

    public double? DefaultTabStopWidth { get; }

    public System.Collections.Generic.IReadOnlyList<PdfTabStop> TabStops { get; }

    internal PdfTableCellParagraph Clone() => new PdfTableCellParagraph(Runs, SpacingAfter, Align, SpacingBefore, LeftIndent, RightIndent, FirstLineIndent, LineHeight, DefaultTabStopWidth, TabStops);
}
