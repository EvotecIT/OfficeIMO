using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Projects this package into a single flat OpenDocument XML tree.</summary>
    public XDocument ToFlatXml() {
        XElement root = new XElement(OdfNamespaces.Office + "document");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", Version.ToToken());
        root.SetAttributeValue(OdfNamespaces.Office + "mimetype", OdfMediaTypes.ForKind(Kind));

        XDocument content = GetXml("content.xml");
        XDocument styles = Package.ContainsEntry("styles.xml") ? GetXml("styles.xml") : OdfPackageTemplates.CreateStyles(Version);
        XDocument meta = Package.ContainsEntry("meta.xml") ? GetXml("meta.xml") : OdfPackageTemplates.CreateMetadata(Version);
        XDocument settings = Package.ContainsEntry("settings.xml") ? GetXml("settings.xml") : OdfPackageTemplates.CreateSettings(Version);
        AddClone(root, meta.Root?.Element(OdfNamespaces.Office + "meta"));
        XElement? flatSettings = settings.Root?.Element(OdfNamespaces.Office + "settings");
        if (flatSettings != null && flatSettings.HasElements) AddClone(root, flatSettings);
        AddClone(root, content.Root?.Element(OdfNamespaces.Office + "scripts"));
        root.Add(MergeContainers(OdfNamespaces.Office + "font-face-decls", content.Root, styles.Root));
        AddClone(root, styles.Root?.Element(OdfNamespaces.Office + "styles"));
        root.Add(MergeContainers(OdfNamespaces.Office + "automatic-styles", content.Root, styles.Root));
        AddClone(root, styles.Root?.Element(OdfNamespaces.Office + "master-styles"));
        XElement body = new XElement(content.Root?.Element(OdfNamespaces.Office + "body")
            ?? throw new InvalidDataException("OpenDocument content has no body."));
        root.Add(body);
        EmbedFlatBinaryData(root);
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    /// <summary>Writes flat OpenDocument XML and returns the serialized bytes with projection diagnostics.</summary>
    public OdfSaveResult SaveFlatXml(Stream destination) {
        XDocument flat = ToFlatXml();
        byte[] bytes = OdfXmlCodec.Save(flat);
        OfficeStreamWriter.WriteAllBytes(destination, bytes);
        return new OdfSaveResult(bytes, CreateFlatXmlSaveReport());
    }

    /// <summary>Writes flat OpenDocument XML to a path and returns the serialized bytes with projection diagnostics.</summary>
    public OdfSaveResult SaveFlatXml(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        XDocument flat = ToFlatXml();
        byte[] bytes = OdfXmlCodec.Save(flat);
        OfficeFileCommit.WriteAllBytes(fullPath, bytes);
        return new OdfSaveResult(bytes, CreateFlatXmlSaveReport());
    }

    /// <summary>Loads a flat ODT, ODS, or ODP XML document.</summary>
    public static OdfDocument LoadFlatXml(Stream stream, OdfLoadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Flat OpenDocument stream must be readable.", nameof(stream));
        OdfLoadOptions effective = (options ?? new OdfLoadOptions()).Normalize();
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, effective.MaxPackageBytes);
        XDocument flat = OdfXmlCodec.Load(bytes, "flat-document.xml", effective.MaxXmlCharacters, effective.MaxXmlDepth);
        return LoadFlatXml(flat, effective);
    }

    /// <summary>Loads a flat ODT, ODS, or ODP XML document from a path.</summary>
    public static OdfDocument LoadFlatXml(string path, OdfLoadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("Flat OpenDocument file does not exist.", fullPath);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return LoadFlatXml(stream, options);
    }

    private static OdfDocument LoadFlatXml(XDocument flat, OdfLoadOptions options) {
        XElement root = flat.Root ?? throw new InvalidDataException("Flat OpenDocument XML has no root element.");
        if (root.Name != OdfNamespaces.Office + "document") throw new InvalidDataException("Flat OpenDocument root must be office:document.");
        string mediaType = (string?)root.Attribute(OdfNamespaces.Office + "mimetype") ?? string.Empty;
        if (!OdfMediaTypes.TryGetKind(mediaType, out OdfDocumentKind kind)) throw new InvalidDataException("Unsupported flat OpenDocument media type '" + mediaType + "'.");
        string? versionToken = (string?)root.Attribute(OdfNamespaces.Office + "version");
        if (!OdfVersionExtensions.TryParse(versionToken, out OdfVersion version)) version = OdfVersion.V1_4;
        OdfPackage package = OdfPackage.Create(kind, version);
        ExtractFlatBinaryData(root, package, options);

        XDocument content = OdfPackageTemplates.CreateContent(kind, version);
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "scripts", root.Element(OdfNamespaces.Office + "scripts"));
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "font-face-decls", root.Element(OdfNamespaces.Office + "font-face-decls"));
        XElement body = new XElement(root.Element(OdfNamespaces.Office + "body")
            ?? throw new InvalidDataException("Flat OpenDocument XML has no office:body."));
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "body", body);

        XDocument styles = OdfPackageTemplates.CreateStyles(version);
        ReplaceContainer(styles.Root!, OdfNamespaces.Office + "styles", root.Element(OdfNamespaces.Office + "styles"));
        ReplaceContainer(styles.Root!, OdfNamespaces.Office + "master-styles", root.Element(OdfNamespaces.Office + "master-styles"));
        SplitFlatAutomaticStyles(root, content.Root!, styles.Root!);

        XDocument meta = OdfPackageTemplates.CreateMetadata(version);
        ReplaceContainer(meta.Root!, OdfNamespaces.Office + "meta", root.Element(OdfNamespaces.Office + "meta"));
        XDocument settings = OdfPackageTemplates.CreateSettings(version);
        XElement? flatSettings = root.Element(OdfNamespaces.Office + "settings");
        ReplaceContainer(settings.Root!, OdfNamespaces.Office + "settings", flatSettings);

        package.AddOrReplaceEntry("content.xml", OdfXmlCodec.Save(content), "text/xml");
        package.AddOrReplaceEntry("styles.xml", OdfXmlCodec.Save(styles), "text/xml");
        package.AddOrReplaceEntry("meta.xml", OdfXmlCodec.Save(meta), "text/xml");
        package.AddOrReplaceEntry("settings.xml", OdfXmlCodec.Save(settings), "text/xml");
        package = OdfPackage.Load(package.Write(new OdfSaveOptions {
            CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
        }), options);
        return CreateForPackage(package, null);
    }

    private void EmbedFlatBinaryData(XElement flatRoot) {
        foreach (XElement image in flatRoot.Descendants(OdfNamespaces.Draw + "image")) {
            string? href = (string?)image.Attribute(OdfNamespaces.XLink + "href");
            if (string.IsNullOrWhiteSpace(href) || href!.Contains("://")) continue;
            string normalized = OdfPackagePath.NormalizeHref(href);
            if (!Package.ContainsEntry(normalized)) continue;
            OdfPackageEntry entry = Package.GetRequiredEntry(normalized);
            image.SetAttributeValue(OdfNamespaces.XLink + "href", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "type", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "show", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "actuate", null);
            image.SetAttributeValue(OdfNamespaces.Draw + "mime-type", entry.MediaType);
            image.Elements(OdfNamespaces.Office + "binary-data").Remove();
            image.Add(new XElement(OdfNamespaces.Office + "binary-data", Convert.ToBase64String(entry.GetOriginalBytes())));
        }
    }

    private static void ExtractFlatBinaryData(XElement flatRoot, OdfPackage package, OdfLoadOptions options) {
        int index = 1;
        foreach (XElement image in flatRoot.Descendants(OdfNamespaces.Draw + "image").ToList()) {
            XElement? binary = image.Element(OdfNamespaces.Office + "binary-data");
            if (binary == null) continue;
            byte[] data;
            try { data = Convert.FromBase64String(new string(binary.Value.Where(character => !char.IsWhiteSpace(character)).ToArray())); }
            catch (FormatException ex) { throw new InvalidDataException("Flat OpenDocument image contains invalid base64 data.", ex); }
            if (data.LongLength > options.MaxEntryUncompressedBytes) throw new InvalidDataException("Flat OpenDocument image exceeds MaxEntryUncompressedBytes.");
            string? mediaType = (string?)image.Attribute(OdfNamespaces.Draw + "mime-type");
            string extension = ImageExtension(mediaType, data);
            if (extension == ".svg") {
                if (!OfficeSvgDrawingReader.TryRead(data, out OfficeDrawing? drawing) || drawing == null) {
                    throw new InvalidDataException("Flat OpenDocument SVG image is not a supported bounded vector image.");
                }
                data = Encoding.UTF8.GetBytes(OfficeDrawingSvgExporter.ToSvg(drawing));
            }
            string path = "Pictures/flat-image" + index++.ToString(CultureInfo.InvariantCulture) + extension;
            package.AddOrReplaceEntry(path, data, ImageMediaType(extension));
            image.SetAttributeValue(OdfNamespaces.XLink + "href", path);
            image.SetAttributeValue(OdfNamespaces.XLink + "type", "simple");
            image.SetAttributeValue(OdfNamespaces.XLink + "show", "embed");
            image.SetAttributeValue(OdfNamespaces.XLink + "actuate", "onLoad");
            binary.Remove();
        }
    }

    private static string DetectImageExtension(byte[] data) {
        if (data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return ".png";
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return ".jpg";
        if (data.Length >= 6 && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46) return ".gif";
        if (data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D) return ".bmp";
        return ".bin";
    }

    private static string ImageExtension(string? mediaType, byte[] data) {
        switch ((mediaType ?? string.Empty).Trim().ToLowerInvariant()) {
            case "image/png": return ".png";
            case "image/jpeg": return ".jpg";
            case "image/gif": return ".gif";
            case "image/svg+xml": return ".svg";
            case "image/bmp": return ".bmp";
            case "image/webp": return ".webp";
            default: return DetectImageExtension(data);
        }
    }

    private static string ImageMediaType(string extension) {
        switch (extension) {
            case ".png": return "image/png";
            case ".jpg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".svg": return "image/svg+xml";
            case ".bmp": return "image/bmp";
            case ".webp": return "image/webp";
            default: return "application/octet-stream";
        }
    }

    private static XElement MergeContainers(XName name, params XElement?[] roots) {
        var result = new XElement(name);
        foreach (XElement? root in roots) {
            XElement? source = root?.Element(name);
            if (source != null) result.Add(source.Nodes().Select(CloneNode));
        }
        return result;
    }

    private static void SplitFlatAutomaticStyles(XElement flatRoot, XElement contentRoot, XElement? stylesRoot) {
        XElement? source = flatRoot.Element(OdfNamespaces.Office + "automatic-styles");
        if (source == null) {
            ReplaceContainer(contentRoot, OdfNamespaces.Office + "automatic-styles", null);
            if (stylesRoot != null) ReplaceContainer(stylesRoot, OdfNamespaces.Office + "automatic-styles", null);
            return;
        }

        XElement[] automaticStyles = source.Elements().ToArray();
        var stylesByName = new Dictionary<string, List<XElement>>(StringComparer.Ordinal);
        foreach (XElement element in automaticStyles) {
            string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
            if (string.IsNullOrEmpty(name)) continue;
            if (!stylesByName.TryGetValue(name!, out List<XElement>? namedStyles)) {
                namedStyles = new List<XElement>();
                stylesByName.Add(name!, namedStyles);
            }
            namedStyles.Add(element);
        }

        var styleScopedElements = new HashSet<XElement>();
        var pendingNames = new Queue<string>();
        var queuedNames = new HashSet<string>(StringComparer.Ordinal);
        void QueueReferences(XElement element) {
            foreach (XAttribute attribute in element.DescendantsAndSelf().Attributes()) {
                if (queuedNames.Add(attribute.Value)) pendingNames.Enqueue(attribute.Value);
            }
        }

        XElement? masters = flatRoot.Element(OdfNamespaces.Office + "master-styles");
        if (masters != null) QueueReferences(masters);
        foreach (XElement element in automaticStyles) {
            if (element.Name != OdfNamespaces.Style + "page-layout"
                && element.Name != OdfNamespaces.Style + "presentation-page-layout") continue;
            if (styleScopedElements.Add(element)) QueueReferences(element);
        }
        while (pendingNames.Count > 0) {
            string name = pendingNames.Dequeue();
            if (!stylesByName.TryGetValue(name, out List<XElement>? namedStyles)) continue;
            foreach (XElement element in namedStyles) {
                if (styleScopedElements.Add(element)) QueueReferences(element);
            }
        }

        var styleScoped = new List<XElement>();
        var contentScoped = new List<XElement>();
        foreach (XElement element in automaticStyles) {
            bool belongsToStyles = styleScopedElements.Contains(element);
            (belongsToStyles ? styleScoped : contentScoped).Add(new XElement(element));
        }

        ReplaceContainer(contentRoot, OdfNamespaces.Office + "automatic-styles",
            new XElement(OdfNamespaces.Office + "automatic-styles", contentScoped));
        if (stylesRoot != null) ReplaceContainer(stylesRoot, OdfNamespaces.Office + "automatic-styles",
            new XElement(OdfNamespaces.Office + "automatic-styles", styleScoped));
    }

    private static void ReplaceContainer(XElement root, XName name, XElement? source) {
        XElement? current = root.Element(name);
        XElement replacement = source == null ? new XElement(name) : new XElement(source);
        if (current == null) root.Add(replacement); else current.ReplaceWith(replacement);
    }

    private static void AddClone(XElement target, XElement? source) { if (source != null) target.Add(new XElement(source)); }
    private static XNode CloneNode(XNode node) => node is XElement element ? new XElement(element) :
        node is XText text ? new XText(text.Value) : node is XComment comment ? new XComment(comment.Value) : new XText(node.ToString());

    private OdfSaveReport CreateFlatXmlSaveReport() {
        var represented = new HashSet<string>(StringComparer.Ordinal) {
            "mimetype", "content.xml", "styles.xml", "meta.xml", "settings.xml", "META-INF/manifest.xml"
        };
        XDocument content = GetXml("content.xml");
        XDocument styles = Package.ContainsEntry("styles.xml") ? GetXml("styles.xml") : OdfPackageTemplates.CreateStyles(Version);
        AddRepresentedImages(content, represented);
        AddRepresentedImages(styles, represented);

        var lossy = Package.Entries.Where(entry => !represented.Contains(entry.Name)).Select(entry => entry.Name).ToList();
        AddUnprojectedPart(lossy, "content.xml", content.Root,
            OdfNamespaces.Office + "scripts", OdfNamespaces.Office + "font-face-decls",
            OdfNamespaces.Office + "automatic-styles", OdfNamespaces.Office + "body");
        AddUnprojectedPart(lossy, "styles.xml", styles.Root,
            OdfNamespaces.Office + "font-face-decls", OdfNamespaces.Office + "styles",
            OdfNamespaces.Office + "automatic-styles", OdfNamespaces.Office + "master-styles");
        if (Package.ContainsEntry("meta.xml")) AddUnprojectedPart(lossy, "meta.xml", GetXml("meta.xml").Root, OdfNamespaces.Office + "meta");
        if (Package.ContainsEntry("settings.xml")) AddUnprojectedPart(lossy, "settings.xml", GetXml("settings.xml").Root, OdfNamespaces.Office + "settings");

        string[] rewritten = represented.Where(path => Package.ContainsEntry(path) && path != "mimetype" && path != "META-INF/manifest.xml")
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();
        return new OdfSaveReport(rewritten, Array.Empty<string>(), Array.Empty<string>(),
            lossy.Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal).ToArray());
    }

    private void AddRepresentedImages(XDocument document, HashSet<string> represented) {
        foreach (XElement image in document.Descendants(OdfNamespaces.Draw + "image")) {
            string? href = (string?)image.Attribute(OdfNamespaces.XLink + "href");
            if (string.IsNullOrWhiteSpace(href) || href!.Contains("://")) continue;
            string normalized = OdfPackagePath.NormalizeHref(href);
            if (Package.ContainsEntry(normalized)) represented.Add(normalized);
        }
    }

    private static void AddUnprojectedPart(List<string> lossy, string partPath, XElement? root, params XName[] projectedChildren) {
        if (root == null) return;
        var projected = new HashSet<XName>(projectedChildren);
        if (root.Elements().Any(element => !projected.Contains(element.Name))) lossy.Add(partPath);
    }

}

public sealed partial class OdtDocument {
    /// <summary>Loads a flat OpenDocument Text XML stream.</summary>
    public new static OdtDocument LoadFlatXml(Stream stream, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(stream, options) as OdtDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Text document.");
    /// <summary>Loads a flat OpenDocument Text XML path.</summary>
    public new static OdtDocument LoadFlatXml(string path, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(path, options) as OdtDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Text document.");
}

public sealed partial class OdsDocument {
    /// <summary>Loads a flat OpenDocument Spreadsheet XML stream.</summary>
    public new static OdsDocument LoadFlatXml(Stream stream, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(stream, options) as OdsDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Spreadsheet document.");
    /// <summary>Loads a flat OpenDocument Spreadsheet XML path.</summary>
    public new static OdsDocument LoadFlatXml(string path, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(path, options) as OdsDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Spreadsheet document.");
}

public sealed partial class OdpPresentation {
    /// <summary>Loads a flat OpenDocument Presentation XML stream.</summary>
    public new static OdpPresentation LoadFlatXml(Stream stream, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(stream, options) as OdpPresentation ?? throw new InvalidDataException("Flat document is not an OpenDocument Presentation document.");
    /// <summary>Loads a flat OpenDocument Presentation XML path.</summary>
    public new static OdpPresentation LoadFlatXml(string path, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(path, options) as OdpPresentation ?? throw new InvalidDataException("Flat document is not an OpenDocument Presentation document.");
}
