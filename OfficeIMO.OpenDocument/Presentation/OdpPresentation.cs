namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Presentation document.</summary>
public sealed partial class OdpPresentation : OdfDocument {
    private OdpPresentation(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Presentation) throw new InvalidDataException("Package is not an OpenDocument Presentation document.");
    }

    /// <summary>Creates an empty ODF 1.4 presentation.</summary>
    public static OdpPresentation Create() => new OdpPresentation(OdfPackage.Create(OdfDocumentKind.Presentation), null);

    /// <summary>Opens an ODP document from a path.</summary>
    public static OdpPresentation Open(string path, OdfOpenOptions? options = null) {
        OdfPackage package = OdfPackage.Open(path, options, out string fullPath);
        return new OdpPresentation(package, fullPath);
    }

    /// <summary>Opens an ODP document from a stream.</summary>
    public static OdpPresentation Open(Stream stream, OdfOpenOptions? options = null) => new OdpPresentation(OdfPackage.Open(stream, options), null);

    internal XElement PresentationBody => GetBody(OdfNamespaces.Office + "presentation");
}
