namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Spreadsheet document.</summary>
public sealed partial class OdsDocument : OdfDocument {
    internal OdsDocument(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Spreadsheet) throw new InvalidDataException("Package is not an OpenDocument Spreadsheet document.");
    }

    /// <summary>Creates an empty ODF 1.4 spreadsheet.</summary>
    public static OdsDocument Create() => new OdsDocument(OdfPackage.Create(OdfDocumentKind.Spreadsheet), null);

    /// <summary>Loads an ODS document from a path.</summary>
    public new static OdsDocument Load(string path, OdfLoadOptions? options = null) {
        OdfPackage package = OdfPackage.Load(path, options, out string fullPath);
        return new OdsDocument(package, fullPath);
    }

    /// <summary>Loads an ODS document from a stream.</summary>
    public new static OdsDocument Load(Stream stream, OdfLoadOptions? options = null) => new OdsDocument(OdfPackage.Load(stream, options), null);

    /// <summary>Asynchronously loads an ODS document from a path.</summary>
    public new static async Task<OdsDocument> LoadAsync(string path, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfDocument document = await OdfDocument.LoadAsync(path, options, cancellationToken).ConfigureAwait(false);
        return document as OdsDocument ?? throw new InvalidDataException("Package is not an OpenDocument Spreadsheet document.");
    }

    /// <summary>Asynchronously loads an ODS document from a caller-owned stream.</summary>
    public new static async Task<OdsDocument> LoadAsync(Stream stream, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfPackage package = await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false);
        return new OdsDocument(package, null);
    }

    internal XElement SpreadsheetBody => GetBody(OdfNamespaces.Office + "spreadsheet");

    /// <summary>Worksheets in document order.</summary>
    public IReadOnlyList<OdsSheet> Sheets => SpreadsheetBody.Elements(OdfNamespaces.Table + "table")
        .Select(element => new OdsSheet(this, element)).ToList();

    /// <summary>Named ranges stored at workbook scope.</summary>
    public IReadOnlyList<OdsNamedRange> NamedRanges => GetNamedExpressions(create: false)?
        .Elements(OdfNamespaces.Table + "named-range")
        .Select(element => new OdsNamedRange(this, element)).ToList() ?? (IReadOnlyList<OdsNamedRange>)Array.Empty<OdsNamedRange>();

    /// <summary>Content validation rules stored at workbook scope.</summary>
    public IReadOnlyList<OdsValidation> Validations => SpreadsheetBody.Element(OdfNamespaces.Table + "content-validations")?
        .Elements(OdfNamespaces.Table + "content-validation")
        .Select(element => new OdsValidation(this, element)).ToList() ?? (IReadOnlyList<OdsValidation>)Array.Empty<OdsValidation>();

    /// <summary>Adds an empty worksheet.</summary>
    public OdsSheet AddSheet(string name) {
        ValidateSheetName(name);
        if (Sheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.Ordinal))) {
            throw new InvalidOperationException($"A worksheet named '{name}' already exists.");
        }
        var element = new XElement(OdfNamespaces.Table + "table",
            new XAttribute(OdfNamespaces.Table + "name", name),
            new XElement(OdfNamespaces.Table + "table-column"),
            new XElement(OdfNamespaces.Table + "table-row", new XElement(OdfNamespaces.Table + "table-cell")));
        SpreadsheetBody.Add(element);
        MarkPartDirty("content.xml");
        return new OdsSheet(this, element);
    }

    /// <summary>Finds a worksheet by its case-sensitive name.</summary>
    public OdsSheet? GetSheet(string name) => Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, name, StringComparison.Ordinal));

    /// <summary>Moves a worksheet to a zero-based position.</summary>
    public void MoveSheet(string name, int newIndex) {
        OdsSheet sheet = GetSheet(name) ?? throw new ArgumentException($"Worksheet '{name}' does not exist.", nameof(name));
        List<XElement> elements = SpreadsheetBody.Elements(OdfNamespaces.Table + "table").ToList();
        if (newIndex < 0 || newIndex >= elements.Count) throw new ArgumentOutOfRangeException(nameof(newIndex));
        XElement moving = sheet.Element;
        moving.Remove();
        elements.Remove(moving);
        if (newIndex >= elements.Count) SpreadsheetBody.Add(moving);
        else elements[newIndex].AddBeforeSelf(moving);
        MarkPartDirty("content.xml");
    }

    /// <summary>Removes a worksheet.</summary>
    public bool RemoveSheet(string name) {
        OdsSheet? sheet = GetSheet(name);
        if (sheet == null) return false;
        sheet.Element.Remove();
        MarkPartDirty("content.xml");
        return true;
    }

    /// <summary>Adds a workbook-scoped named range using ODF address syntax.</summary>
    public OdsNamedRange AddNamedRange(string name, string cellRangeAddress, string? baseCellAddress = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Named range name cannot be empty.", nameof(name));
        if (string.IsNullOrWhiteSpace(cellRangeAddress)) throw new ArgumentException("Cell range address cannot be empty.", nameof(cellRangeAddress));
        XElement container = GetNamedExpressions(create: true)!;
        if (container.Elements(OdfNamespaces.Table + "named-range").Any(element =>
                string.Equals((string?)element.Attribute(OdfNamespaces.Table + "name"), name, StringComparison.Ordinal))) {
            throw new InvalidOperationException($"A named range called '{name}' already exists.");
        }
        var element = new XElement(OdfNamespaces.Table + "named-range",
            new XAttribute(OdfNamespaces.Table + "name", name),
            new XAttribute(OdfNamespaces.Table + "cell-range-address", cellRangeAddress),
            new XAttribute(OdfNamespaces.Table + "base-cell-address", baseCellAddress ?? cellRangeAddress.Split(':')[0]));
        container.Add(element);
        MarkPartDirty("content.xml");
        return new OdsNamedRange(this, element);
    }

    /// <summary>Adds a validation rule with a preserved ODF condition expression.</summary>
    public OdsValidation AddValidation(string name, string condition, bool allowEmptyCell = true) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Validation name cannot be empty.", nameof(name));
        if (string.IsNullOrWhiteSpace(condition)) throw new ArgumentException("Validation condition cannot be empty.", nameof(condition));
        XElement? container = SpreadsheetBody.Element(OdfNamespaces.Table + "content-validations");
        if (container == null) {
            container = new XElement(OdfNamespaces.Table + "content-validations");
            SpreadsheetBody.AddFirst(container);
        }
        var element = new XElement(OdfNamespaces.Table + "content-validation",
            new XAttribute(OdfNamespaces.Table + "name", name),
            new XAttribute(OdfNamespaces.Table + "condition", condition),
            new XAttribute(OdfNamespaces.Table + "allow-empty-cell", allowEmptyCell ? "true" : "false"));
        container.Add(element);
        MarkPartDirty("content.xml");
        return new OdsValidation(this, element);
    }

    private XElement? GetNamedExpressions(bool create) {
        XElement? element = SpreadsheetBody.Element(OdfNamespaces.Table + "named-expressions");
        if (element == null && create) {
            element = new XElement(OdfNamespaces.Table + "named-expressions");
            SpreadsheetBody.Add(element);
            MarkPartDirty("content.xml");
        }
        return element;
    }

    private static void ValidateSheetName(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Worksheet name cannot be empty.", nameof(name));
    }
}
