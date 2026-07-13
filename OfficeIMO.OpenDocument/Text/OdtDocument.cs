namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Text document.</summary>
public sealed partial class OdtDocument : OdfDocument {
    internal OdtDocument(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Text) throw new InvalidDataException("Package is not an OpenDocument Text document.");
    }

    /// <summary>Creates an empty ODF 1.4 text document.</summary>
    public static OdtDocument Create() => new OdtDocument(OdfPackage.Create(OdfDocumentKind.Text), null);

    /// <summary>Loads an ODT document from a path.</summary>
    public new static OdtDocument Load(string path, OdfLoadOptions? options = null) {
        OdfPackage package = OdfPackage.Load(path, options, out string fullPath);
        return new OdtDocument(package, fullPath);
    }

    /// <summary>Loads an ODT document from a stream.</summary>
    public new static OdtDocument Load(Stream stream, OdfLoadOptions? options = null) => new OdtDocument(OdfPackage.Load(stream, options), null);

    /// <summary>Asynchronously loads an ODT document from a path.</summary>
    public new static async Task<OdtDocument> LoadAsync(string path, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfDocument document = await OdfDocument.LoadAsync(path, options, cancellationToken).ConfigureAwait(false);
        return document as OdtDocument ?? throw new InvalidDataException("Package is not an OpenDocument Text document.");
    }

    /// <summary>Asynchronously loads an ODT document from a caller-owned stream.</summary>
    public new static async Task<OdtDocument> LoadAsync(Stream stream, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfPackage package = await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false);
        return new OdtDocument(package, null);
    }

    internal XElement TextBody => GetBody(OdfNamespaces.Office + "text");

    /// <summary>Page layout, margins, header, and footer through the first master page.</summary>
    public OdtPageLayout PageLayout => OdtPageLayout.GetOrCreate(this);
}
