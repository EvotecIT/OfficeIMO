namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Text document.</summary>
public sealed partial class OdtDocument : OdfDocument {
    internal OdtDocument(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Text) throw new InvalidDataException("Package is not an OpenDocument Text document.");
    }

    /// <summary>Creates an empty ODF 1.4 text document.</summary>
    public static OdtDocument Create() => new OdtDocument(OdfPackage.Create(OdfDocumentKind.Text), null);

    /// <summary>Opens an ODT document from a path.</summary>
    public static OdtDocument Open(string path, OdfOpenOptions? options = null) {
        OdfPackage package = OdfPackage.Open(path, options, out string fullPath);
        return new OdtDocument(package, fullPath);
    }

    /// <summary>Opens an ODT document from a stream.</summary>
    public static OdtDocument Open(Stream stream, OdfOpenOptions? options = null) => new OdtDocument(OdfPackage.Open(stream, options), null);

    internal XElement TextBody => GetBody(OdfNamespaces.Office + "text");

    /// <summary>Page layout, margins, header, and footer through the first master page.</summary>
    public OdtPageLayout PageLayout => OdtPageLayout.GetOrCreate(this);
}
