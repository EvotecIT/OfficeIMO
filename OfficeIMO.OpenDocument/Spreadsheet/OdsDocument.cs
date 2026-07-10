namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Spreadsheet document.</summary>
public sealed partial class OdsDocument : OdfDocument {
    private OdsDocument(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Spreadsheet) throw new InvalidDataException("Package is not an OpenDocument Spreadsheet document.");
    }

    /// <summary>Creates an empty ODF 1.4 spreadsheet.</summary>
    public static OdsDocument Create() => new OdsDocument(OdfPackage.Create(OdfDocumentKind.Spreadsheet), null);

    /// <summary>Opens an ODS document from a path.</summary>
    public static OdsDocument Open(string path, OdfOpenOptions? options = null) {
        OdfPackage package = OdfPackage.Open(path, options, out string fullPath);
        return new OdsDocument(package, fullPath);
    }

    /// <summary>Opens an ODS document from a stream.</summary>
    public static OdsDocument Open(Stream stream, OdfOpenOptions? options = null) => new OdsDocument(OdfPackage.Open(stream, options), null);

    internal XElement SpreadsheetBody => GetBody(OdfNamespaces.Office + "spreadsheet");
}
