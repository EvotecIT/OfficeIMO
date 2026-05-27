namespace OfficeIMO.Pdf;

/// <summary>
/// Logical PDF element categories exposed by the first-party read model.
/// </summary>
public enum PdfLogicalElementKind {
    /// <summary>Line-level text recovered from positioned PDF text spans.</summary>
    TextBlock,
    /// <summary>Detected bullet or numbered list item.</summary>
    ListItem,
    /// <summary>Detected leader row such as label plus dotted value.</summary>
    LeaderRow,
    /// <summary>Detected table-like region.</summary>
    Table,
    /// <summary>Image XObject referenced by the page.</summary>
    Image
}

/// <summary>
/// Common shape for logical page elements extracted from a PDF page.
/// </summary>
public interface IPdfLogicalElement {
    /// <summary>One-based source page number.</summary>
    int PageNumber { get; }

    /// <summary>Element kind.</summary>
    PdfLogicalElementKind Kind { get; }
}

/// <summary>
/// First-party logical read model for a parser-supported PDF.
/// </summary>
public sealed class PdfLogicalDocument {
    private IReadOnlyList<IPdfLogicalElement>? _elements;

    private PdfLogicalDocument(PdfMetadata metadata, IReadOnlyList<PdfLogicalPage> pages, IReadOnlyList<PdfFormField> formFields) {
        Metadata = metadata;
        Pages = pages;
        FormFields = formFields;
    }

    /// <summary>Document metadata read from the PDF Info dictionary when available.</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Logical pages in document order.</summary>
    public IReadOnlyList<PdfLogicalPage> Pages { get; }

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>Number of pages in the logical document.</summary>
    public int PageCount => Pages.Count;

    /// <summary>True when at least one simple AcroForm field was read from the document catalog.</summary>
    public bool HasFormFields => FormFields.Count > 0;

    /// <summary>All logical page elements flattened in page order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements {
        get {
            if (_elements is not null) {
                return _elements;
            }

            var elements = new List<IPdfLogicalElement>();
            for (int i = 0; i < Pages.Count; i++) {
                elements.AddRange(Pages[i].Elements);
            }

            _elements = elements.AsReadOnly();
            return _elements;
        }
    }

    /// <summary>Loads a PDF from bytes and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return From(PdfReadDocument.Load(pdf), options);
    }

    /// <summary>Loads a PDF from a file path and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return From(PdfReadDocument.Load(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(Stream stream, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        return From(PdfReadDocument.Load(stream), options);
    }

    /// <summary>Builds the logical read model from an already parsed PDF document.</summary>
    public static PdfLogicalDocument From(PdfReadDocument document, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(document, nameof(document));

        var pages = new List<PdfLogicalPage>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(PdfLogicalPage.From(document.Pages[i], i + 1, options));
        }

        return new PdfLogicalDocument(document.Metadata, pages.AsReadOnly(), document.FormFields);
    }
}

/// <summary>
/// Logical view of a single PDF page.
/// </summary>
public sealed class PdfLogicalPage {
    private PdfLogicalPage(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        IReadOnlyList<IPdfLogicalElement> elements,
        IReadOnlyList<PdfLogicalTextBlock> textBlocks,
        IReadOnlyList<PdfLogicalTable> tables,
        IReadOnlyList<PdfLogicalImage> images,
        IReadOnlyList<PdfLinkAnnotation> linkAnnotations) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        Elements = elements;
        TextBlocks = textBlocks;
        Tables = tables;
        Images = images;
        LinkAnnotations = linkAnnotations;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Inherited page rotation normalized to 0, 90, 180, or 270.</summary>
    public int RotationDegrees { get; }

    /// <summary>Logical elements in extraction order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements { get; }

    /// <summary>Line-level text blocks extracted from positioned text spans.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks { get; }

    /// <summary>Detected table-like regions.</summary>
    public IReadOnlyList<PdfLogicalTable> Tables { get; }

    /// <summary>Image XObjects referenced by the page.</summary>
    public IReadOnlyList<PdfLogicalImage> Images { get; }

    /// <summary>Simple link annotations read from the page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations { get; }

    internal static PdfLogicalPage From(PdfReadPage page, int pageNumber, PdfTextLayoutOptions? options) {
        var size = page.GetPageSize();
        var structured = page.ExtractStructured(options);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var listLines = new HashSet<string>(structured.ListItems.Select(NormalizeForKindComparison), StringComparer.Ordinal);

        foreach (var line in structured.LinesDetailed) {
            string text = line.Text?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            var kind = listLines.Contains(NormalizeForKindComparison(text)) || LooksLikeListItem(text)
                ? PdfLogicalElementKind.ListItem
                : PdfLogicalElementKind.TextBlock;
            var block = new PdfLogicalTextBlock(pageNumber, kind, text, line.XStart, line.XEnd, line.Y, line.SpanCount);
            textBlocks.Add(block);
            elements.Add(block);
        }

        foreach (var row in structured.LeaderRows) {
            if (row.Length < 2) {
                continue;
            }

            var leader = new PdfLogicalLeaderRow(pageNumber, row[0], row[1]);
            elements.Add(leader);
        }

        foreach (var table in structured.TablesDetailed) {
            var logicalTable = PdfLogicalTable.From(pageNumber, table);
            tables.Add(logicalTable);
            elements.Add(logicalTable);
        }

        foreach (var image in page.GetImages(pageNumber)) {
            var logicalImage = new PdfLogicalImage(image);
            images.Add(logicalImage);
            elements.Add(logicalImage);
        }

        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            tables.AsReadOnly(),
            images.AsReadOnly(),
            page.GetLinkAnnotations());
    }

    private static string NormalizeForKindComparison(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            if (!char.IsWhiteSpace(text[i])) {
                builder.Append(text[i]);
            }
        }

        return builder.ToString();
    }

    private static bool LooksLikeListItem(string text) {
        string trimmed = text.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        char marker = trimmed[0];
        if (marker == '\u2022' || marker == '-' || marker == '*' || marker == '\u25CF') {
            return true;
        }

        int index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        return index > 0 &&
            index < trimmed.Length &&
            (trimmed[index] == '.' || trimmed[index] == ')');
    }
}

/// <summary>
/// Line-level text block extracted from a PDF page.
/// </summary>
public sealed class PdfLogicalTextBlock : IPdfLogicalElement {
    internal PdfLogicalTextBlock(int pageNumber, PdfLogicalElementKind kind, string text, double xStart, double xEnd, double baselineY, int spanCount) {
        PageNumber = pageNumber;
        Kind = kind;
        Text = text;
        XStart = xStart;
        XEnd = xEnd;
        BaselineY = baselineY;
        SpanCount = spanCount;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind { get; }

    /// <summary>Extracted text for the line-level block.</summary>
    public string Text { get; }

    /// <summary>Leftmost X coordinate in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Rightmost X coordinate in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Baseline Y coordinate in PDF points from the bottom of the page.</summary>
    public double BaselineY { get; }

    /// <summary>Number of text spans merged into this block.</summary>
    public int SpanCount { get; }
}

/// <summary>
/// Detected leader row such as a table-of-contents or label/value row.
/// </summary>
public sealed class PdfLogicalLeaderRow : IPdfLogicalElement {
    internal PdfLogicalLeaderRow(int pageNumber, string label, string value) {
        PageNumber = pageNumber;
        Label = label;
        Value = value;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.LeaderRow;

    /// <summary>Leader row label.</summary>
    public string Label { get; }

    /// <summary>Leader row trailing value.</summary>
    public string Value { get; }
}

/// <summary>
/// Detected table-like region with simple geometry.
/// </summary>
public sealed class PdfLogicalTable : IPdfLogicalElement {
    private PdfLogicalTable(
        int pageNumber,
        string kind,
        double yTop,
        double yBottom,
        IReadOnlyList<PdfLogicalTableColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> rows) {
        PageNumber = pageNumber;
        DetectionKind = kind;
        YTop = yTop;
        YBottom = yBottom;
        Columns = columns;
        Rows = rows;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Table;

    /// <summary>Detection heuristic that produced the table.</summary>
    public string DetectionKind { get; }

    /// <summary>Top Y coordinate of the detected table band.</summary>
    public double YTop { get; }

    /// <summary>Bottom Y coordinate of the detected table band.</summary>
    public double YBottom { get; }

    /// <summary>Detected table columns.</summary>
    public IReadOnlyList<PdfLogicalTableColumn> Columns { get; }

    /// <summary>Extracted table rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    internal static PdfLogicalTable From(int pageNumber, StructuredTable table) {
        var columns = new List<PdfLogicalTableColumn>(table.Columns.Count);
        for (int i = 0; i < table.Columns.Count; i++) {
            columns.Add(new PdfLogicalTableColumn(table.Columns[i].From, table.Columns[i].To));
        }

        var rows = new List<IReadOnlyList<string>>(table.Rows.Count);
        for (int i = 0; i < table.Rows.Count; i++) {
            rows.Add(Array.AsReadOnly((string[])table.Rows[i].Clone()));
        }

        return new PdfLogicalTable(
            pageNumber,
            table.Kind,
            table.YTop,
            table.YBottom,
            columns.AsReadOnly(),
            rows.AsReadOnly());
    }
}

/// <summary>
/// Detected table column geometry.
/// </summary>
public sealed class PdfLogicalTableColumn {
    internal PdfLogicalTableColumn(double from, double to) {
        From = from;
        To = to;
    }

    /// <summary>Left X coordinate in PDF points.</summary>
    public double From { get; }

    /// <summary>Right X coordinate in PDF points.</summary>
    public double To { get; }
}

/// <summary>
/// Image XObject entry in the logical page model.
/// </summary>
public sealed class PdfLogicalImage : IPdfLogicalElement {
    internal PdfLogicalImage(PdfExtractedImage image) {
        SourceImage = image;
    }

    /// <inheritdoc />
    public int PageNumber => SourceImage.PageNumber;

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Image;

    /// <summary>Underlying extracted image payload and metadata.</summary>
    public PdfExtractedImage SourceImage { get; }

    /// <summary>PDF image resource name.</summary>
    public string ResourceName => SourceImage.ResourceName;

    /// <summary>Image width in pixels.</summary>
    public int Width => SourceImage.Width;

    /// <summary>Image height in pixels.</summary>
    public int Height => SourceImage.Height;

    /// <summary>Suggested MIME type when bytes are a complete image file.</summary>
    public string? MimeType => SourceImage.MimeType;
}
