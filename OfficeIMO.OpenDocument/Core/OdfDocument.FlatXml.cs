namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Projects this package into a single flat OpenDocument XML tree.</summary>
    public XDocument ToFlatXml() {
        ThrowIfDisposed();
        XElement root = new XElement(OdfNamespaces.Office + "document");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(OdfNamespaces.Office + "version", Version.ToToken());
        root.SetAttributeValue(OdfNamespaces.Office + "mimetype", OdfMediaTypes.ForKind(Kind));

        XDocument content = GetXml("content.xml");
        XDocument styles = GetXml("styles.xml");
        XDocument meta = GetXml("meta.xml");
        XDocument settings = GetXml("settings.xml");
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
        EmbedFlatBinaryData(body);
        root.Add(body);
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    /// <summary>Writes flat OpenDocument XML without closing the destination stream.</summary>
    public void SaveFlatXml(Stream destination) {
        ThrowIfDisposed();
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        XDocument flat = ToFlatXml();
        byte[] bytes = OdfXmlCodec.Save(flat);
        destination.Write(bytes, 0, bytes.Length);
        LastSaveReport = CreateFlatXmlSaveReport();
    }

    /// <summary>Writes flat OpenDocument XML to a path.</summary>
    public void SaveFlatXml(string path) {
        ThrowIfDisposed();
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            XDocument flat = ToFlatXml();
            File.WriteAllBytes(temporary, OdfXmlCodec.Save(flat));
            ReplaceFile(temporary, fullPath);
            LastSaveReport = CreateFlatXmlSaveReport();
        } finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    /// <summary>Opens a flat ODT, ODS, or ODP XML document.</summary>
    public static OdfDocument OpenFlatXml(Stream stream, OdfOpenOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Flat OpenDocument stream must be readable.", nameof(stream));
        OdfOpenOptions effective = (options ?? new OdfOpenOptions()).Normalize();
        byte[] bytes = ReadFlatBytes(stream, effective.MaxPackageBytes);
        XDocument flat = OdfXmlCodec.Load(bytes, "flat-document.xml", effective.MaxXmlCharacters, effective.MaxXmlDepth);
        return OpenFlatXml(flat, effective);
    }

    /// <summary>Opens a flat ODT, ODS, or ODP XML document from a path.</summary>
    public static OdfDocument OpenFlatXml(string path, OdfOpenOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("Flat OpenDocument file does not exist.", fullPath);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return OpenFlatXml(stream, options);
    }

    private static OdfDocument OpenFlatXml(XDocument flat, OdfOpenOptions options) {
        XElement root = flat.Root ?? throw new InvalidDataException("Flat OpenDocument XML has no root element.");
        if (root.Name != OdfNamespaces.Office + "document") throw new InvalidDataException("Flat OpenDocument root must be office:document.");
        string mediaType = (string?)root.Attribute(OdfNamespaces.Office + "mimetype") ?? string.Empty;
        if (!OdfMediaTypes.TryGetKind(mediaType, out OdfDocumentKind kind)) throw new InvalidDataException("Unsupported flat OpenDocument media type '" + mediaType + "'.");
        string? versionToken = (string?)root.Attribute(OdfNamespaces.Office + "version");
        if (!OdfVersionExtensions.TryParse(versionToken, out OdfVersion version)) version = OdfVersion.V1_4;
        OdfPackage package = OdfPackage.Create(kind, version);

        XDocument content = OdfPackageTemplates.CreateContent(kind, version);
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "scripts", root.Element(OdfNamespaces.Office + "scripts"));
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "font-face-decls", root.Element(OdfNamespaces.Office + "font-face-decls"));
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "automatic-styles", root.Element(OdfNamespaces.Office + "automatic-styles"));
        XElement body = new XElement(root.Element(OdfNamespaces.Office + "body")
            ?? throw new InvalidDataException("Flat OpenDocument XML has no office:body."));
        ExtractFlatBinaryData(body, package, options);
        ReplaceContainer(content.Root!, OdfNamespaces.Office + "body", body);

        XDocument styles = OdfPackageTemplates.CreateStyles(version);
        ReplaceContainer(styles.Root!, OdfNamespaces.Office + "styles", root.Element(OdfNamespaces.Office + "styles"));
        ReplaceContainer(styles.Root!, OdfNamespaces.Office + "master-styles", root.Element(OdfNamespaces.Office + "master-styles"));

        XDocument meta = OdfPackageTemplates.CreateMetadata(version);
        ReplaceContainer(meta.Root!, OdfNamespaces.Office + "meta", root.Element(OdfNamespaces.Office + "meta"));
        XDocument settings = OdfPackageTemplates.CreateSettings(version);
        XElement? flatSettings = root.Element(OdfNamespaces.Office + "settings");
        ReplaceContainer(settings.Root!, OdfNamespaces.Office + "settings", flatSettings);

        package.AddOrReplaceEntry("content.xml", OdfXmlCodec.Save(content), "text/xml");
        package.AddOrReplaceEntry("styles.xml", OdfXmlCodec.Save(styles), "text/xml");
        package.AddOrReplaceEntry("meta.xml", OdfXmlCodec.Save(meta), "text/xml");
        package.AddOrReplaceEntry("settings.xml", OdfXmlCodec.Save(settings), "text/xml");
        package = OdfPackage.Open(package.Write(), options);
        return CreateForPackage(package, null);
    }

    private void EmbedFlatBinaryData(XElement body) {
        foreach (XElement image in body.Descendants(OdfNamespaces.Draw + "image")) {
            string? href = (string?)image.Attribute(OdfNamespaces.XLink + "href");
            if (string.IsNullOrWhiteSpace(href) || href!.Contains("://") || !Package.ContainsEntry(href)) continue;
            OdfPackageEntry entry = Package.GetRequiredEntry(href);
            image.SetAttributeValue(OdfNamespaces.XLink + "href", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "type", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "show", null);
            image.SetAttributeValue(OdfNamespaces.XLink + "actuate", null);
            image.SetAttributeValue(OdfNamespaces.Draw + "mime-type", entry.MediaType);
            image.Elements(OdfNamespaces.Office + "binary-data").Remove();
            image.Add(new XElement(OdfNamespaces.Office + "binary-data", Convert.ToBase64String(entry.GetOriginalBytes())));
        }
    }

    private static void ExtractFlatBinaryData(XElement body, OdfPackage package, OdfOpenOptions options) {
        int index = 1;
        foreach (XElement image in body.Descendants(OdfNamespaces.Draw + "image").ToList()) {
            XElement? binary = image.Element(OdfNamespaces.Office + "binary-data");
            if (binary == null) continue;
            byte[] data;
            try { data = Convert.FromBase64String(new string(binary.Value.Where(character => !char.IsWhiteSpace(character)).ToArray())); }
            catch (FormatException ex) { throw new InvalidDataException("Flat OpenDocument image contains invalid base64 data.", ex); }
            if (data.LongLength > options.MaxEntryUncompressedBytes) throw new InvalidDataException("Flat OpenDocument image exceeds MaxEntryUncompressedBytes.");
            string extension = DetectImageExtension(data);
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

    private static string ImageMediaType(string extension) {
        switch (extension) {
            case ".png": return "image/png";
            case ".jpg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".bmp": return "image/bmp";
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
        foreach (XElement image in content.Descendants(OdfNamespaces.Draw + "image")) {
            string? href = (string?)image.Attribute(OdfNamespaces.XLink + "href");
            if (!string.IsNullOrWhiteSpace(href) && !href!.Contains("://") && Package.ContainsEntry(href)) represented.Add(href);
        }

        var lossy = Package.Entries.Where(entry => !represented.Contains(entry.Name)).Select(entry => entry.Name).ToList();
        AddUnprojectedPart(lossy, "content.xml", content.Root,
            OdfNamespaces.Office + "scripts", OdfNamespaces.Office + "font-face-decls",
            OdfNamespaces.Office + "automatic-styles", OdfNamespaces.Office + "body");
        XDocument styles = GetXml("styles.xml");
        AddUnprojectedPart(lossy, "styles.xml", styles.Root,
            OdfNamespaces.Office + "font-face-decls", OdfNamespaces.Office + "styles",
            OdfNamespaces.Office + "automatic-styles", OdfNamespaces.Office + "master-styles");
        AddUnprojectedPart(lossy, "meta.xml", GetXml("meta.xml").Root, OdfNamespaces.Office + "meta");
        AddUnprojectedPart(lossy, "settings.xml", GetXml("settings.xml").Root, OdfNamespaces.Office + "settings");

        string[] rewritten = represented.Where(path => Package.ContainsEntry(path) && path != "mimetype" && path != "META-INF/manifest.xml")
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();
        return new OdfSaveReport(rewritten, Array.Empty<string>(), Array.Empty<string>(),
            lossy.Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal).ToArray());
    }

    private static void AddUnprojectedPart(List<string> lossy, string partPath, XElement? root, params XName[] projectedChildren) {
        if (root == null) return;
        var projected = new HashSet<XName>(projectedChildren);
        if (root.Elements().Any(element => !projected.Contains(element.Name))) lossy.Add(partPath);
    }

    private static byte[] ReadFlatBytes(Stream stream, long maxBytes) {
        if (stream.CanSeek && stream.Length - stream.Position > maxBytes) throw new InvalidDataException("Flat OpenDocument stream exceeds MaxPackageBytes.");
        using var output = new MemoryStream();
        var buffer = new byte[81920]; long total = 0; int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            total += read;
            if (total > maxBytes) throw new InvalidDataException("Flat OpenDocument stream exceeds MaxPackageBytes.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }
}

public sealed partial class OdtDocument {
    /// <summary>Opens a flat OpenDocument Text XML stream.</summary>
    public new static OdtDocument OpenFlatXml(Stream stream, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(stream, options) as OdtDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Text document.");
    /// <summary>Opens a flat OpenDocument Text XML path.</summary>
    public new static OdtDocument OpenFlatXml(string path, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(path, options) as OdtDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Text document.");
}

public sealed partial class OdsDocument {
    /// <summary>Opens a flat OpenDocument Spreadsheet XML stream.</summary>
    public new static OdsDocument OpenFlatXml(Stream stream, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(stream, options) as OdsDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Spreadsheet document.");
    /// <summary>Opens a flat OpenDocument Spreadsheet XML path.</summary>
    public new static OdsDocument OpenFlatXml(string path, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(path, options) as OdsDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument Spreadsheet document.");
}

public sealed partial class OdpPresentation {
    /// <summary>Opens a flat OpenDocument Presentation XML stream.</summary>
    public new static OdpPresentation OpenFlatXml(Stream stream, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(stream, options) as OdpPresentation ?? throw new InvalidDataException("Flat document is not an OpenDocument Presentation document.");
    /// <summary>Opens a flat OpenDocument Presentation XML path.</summary>
    public new static OdpPresentation OpenFlatXml(string path, OdfOpenOptions? options = null) =>
        OdfDocument.OpenFlatXml(path, options) as OdpPresentation ?? throw new InvalidDataException("Flat document is not an OpenDocument Presentation document.");
}
