namespace OfficeIMO.OpenDocument;

internal static class OdfPackageTemplates {
    internal static XDocument CreateContent(OdfDocumentKind kind, OdfVersion version) {
        XElement root = new XElement(OdfNamespaces.Office + "document-content");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", version.ToToken());
        root.Add(new XElement(OdfNamespaces.Office + "scripts"));
        root.Add(new XElement(OdfNamespaces.Office + "font-face-decls"));
        root.Add(new XElement(OdfNamespaces.Office + "automatic-styles"));

        XName bodyName;
        switch (kind) {
            case OdfDocumentKind.Text: bodyName = OdfNamespaces.Office + "text"; break;
            case OdfDocumentKind.Spreadsheet: bodyName = OdfNamespaces.Office + "spreadsheet"; break;
            case OdfDocumentKind.Presentation: bodyName = OdfNamespaces.Office + "presentation"; break;
            default: throw new ArgumentOutOfRangeException(nameof(kind));
        }
        root.Add(new XElement(OdfNamespaces.Office + "body", new XElement(bodyName)));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    internal static XDocument CreateStyles(OdfVersion version) {
        XElement root = new XElement(OdfNamespaces.Office + "document-styles");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", version.ToToken());
        root.Add(new XElement(OdfNamespaces.Office + "font-face-decls"));
        root.Add(new XElement(OdfNamespaces.Office + "styles"));
        root.Add(new XElement(OdfNamespaces.Office + "automatic-styles"));
        root.Add(new XElement(OdfNamespaces.Office + "master-styles"));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    internal static XDocument CreateMetadata(OdfVersion version) {
        XElement root = new XElement(OdfNamespaces.Office + "document-meta");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", version.ToToken());
        root.Add(new XElement(OdfNamespaces.Office + "meta",
            new XElement(OdfNamespaces.Meta + "generator", "OfficeIMO.OpenDocument")));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    internal static XDocument CreateSettings(OdfVersion version) {
        XElement root = new XElement(OdfNamespaces.Office + "document-settings");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", version.ToToken());
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    internal static XDocument CreateManifest(OdfDocumentKind kind, OdfVersion version) {
        XElement root = new XElement(OdfNamespaces.Manifest + "manifest",
            new XAttribute(XNamespace.Xmlns + "manifest", OdfNamespaces.Manifest.NamespaceName),
            new XAttribute(OdfNamespaces.Manifest + "version", version.ToToken()),
            FileEntry("/", OdfMediaTypes.ForKind(kind), version.ToToken()),
            FileEntry("content.xml", "text/xml", null),
            FileEntry("styles.xml", "text/xml", null),
            FileEntry("meta.xml", "text/xml", null),
            FileEntry("settings.xml", "text/xml", null));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    internal static XElement FileEntry(string path, string mediaType, string? version) {
        var element = new XElement(OdfNamespaces.Manifest + "file-entry",
            new XAttribute(OdfNamespaces.Manifest + "full-path", path),
            new XAttribute(OdfNamespaces.Manifest + "media-type", mediaType));
        if (!string.IsNullOrEmpty(version)) {
            element.SetAttributeValue(OdfNamespaces.Manifest + "version", version);
        }
        return element;
    }
}
