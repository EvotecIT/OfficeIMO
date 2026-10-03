namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>Adds a local manifest resource at a canonical container path. Payloads are copied and never executed.</summary>
    public EpubManifestItem AddResource(string id, string containerPath, string mediaType, byte[] data, string? properties = null) {
        XElement element = PrepareResourceDeclaration(id, containerPath, mediaType, data, properties);
        byte[] payload = (byte[])data.Clone();
        EditPackageElement(RequireSection("manifest"), proposed => proposed.Add(new XElement(element)), data.LongLength);
        _entries.Add(containerPath, payload);
        _retainedBytes += data.LongLength;
        return RequireManifestItem(id);
    }

    private XElement PrepareResourceDeclaration(string id, string containerPath, string mediaType, byte[] data, string? properties) {
        VerifyAvailableId(id); RequireText(mediaType, nameof(mediaType));
        VerifyPropertiesVersion(properties);
        if (data == null) throw new ArgumentNullException(nameof(data));
        string path = VerifyContentPath(containerPath);
        if (_entries.ContainsKey(path) || Manifest.Any(item => item.Reference.ContainerPath == path)) throw new ArgumentException("Entry path already exists: " + path);
        EnsureEntryBudget(1);
        if (_entries.Keys.Any(existing => string.Equals(existing.Normalize(NormalizationForm.FormC), path.Normalize(NormalizationForm.FormC), StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Resource paths must remain distinct after Unicode normalization and case comparison.", nameof(containerPath));
        if (IsScriptMediaType(mediaType) || HasToken(properties, "scripted"))
            throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
        if (data.LongLength > _maximumEntryBytes) throw new InvalidDataException("Resource exceeds the publication's retained entry-byte limit.");
        var element = new XElement(Opf + "item", new XAttribute("id", id), new XAttribute("href", RelativeHref(PackagePath, path)),
            new XAttribute("media-type", mediaType));
        element.SetAttributeValue("properties", NormalizeProperties(properties));
        return element;
    }

    /// <summary>Replaces one local resource payload. Encrypted/obfuscated bytes require an explicit re-keying implementation.</summary>
    public void UpdateResource(string manifestId, byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        EpubManifestItem item = RequireManifestItem(manifestId);
        string path = RequireLocalPath(item);
        ReplaceResourcePayload(path, data);
    }

    private void ReplaceResourcePayload(string path, byte[] data) {
        EnsureResourceMutationAllowed(path);
        if (_encryption.Any(encryption => encryption.Path == path)) throw new NotSupportedException("Encrypted/obfuscated resource replacement is unsupported.");
        EnsureEntryBudget(_entries.ContainsKey(path) ? 0 : 1);
        long previousLength = _entries.TryGetValue(path, out byte[]? previous) ? previous.LongLength : 0;
        EnsurePayloadBudget(data, previousLength);
        _entries[path] = (byte[])data.Clone();
        _retainedBytes += data.LongLength - previousLength;
        MarkChanged();
    }

    /// <summary>Returns an independent copy of retained resource bytes, including original font-obfuscated bytes.</summary>
    public byte[] GetResourceBytes(string manifestId) {
        string path = RequireLocalPath(RequireManifestItem(manifestId));
        if (!_entries.TryGetValue(path, out byte[]? data)) throw new InvalidDataException("Resource is missing: " + path);
        return (byte[])data.Clone();
    }

    /// <summary>Returns an independent XML content document for targeted editing.</summary>
    public XDocument GetContentXml(string manifestId) => ParseXml(GetResourceBytes(manifestId), _maximumEntryBytes);

    /// <summary>Replaces content XML after validating a supported non-scripted XHTML or SVG root.</summary>
    public void SetContentXml(string manifestId, XDocument content) {
        if (content == null) throw new ArgumentNullException(nameof(content));
        ValidateContent(content, RequireManifestItem(manifestId).MediaType);
        UpdateResource(manifestId, SerializeXml(content, _maximumEntryBytes));
    }

    /// <summary>Adds a well-formed XHTML body fragment as a chapter, with a linear spine position and a TOC entry.</summary>
    public EpubManifestItem AddChapter(string id, string containerPath, string title, string xhtmlBody,
        IEnumerable<string>? stylesheets = null, bool linear = true) {
        RequireText(title, nameof(title));
        VerifyContentPath(containerPath);
        if (xhtmlBody == null) throw new ArgumentNullException(nameof(xhtmlBody));
        XElement body;
        using (var reader = XmlReader.Create(new StringReader("<body xmlns='" + Html.NamespaceName + "' xmlns:epub='" + Ops.NamespaceName + "'>" + xhtmlBody + "</body>"),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = _maximumEntryBytes })) {
            body = XElement.Load(reader, LoadOptions.PreserveWhitespace);
        }
        var head = new XElement(Html + "head", new XElement(Html + "title", title));
        foreach (string stylesheet in stylesheets ?? Array.Empty<string>()) {
            EpubManifestItem style = RequireManifestItem(stylesheet);
            if (!HasMediaType(style.MediaType, "text/css")) throw new ArgumentException("Stylesheet id must select a CSS resource.", nameof(stylesheets));
            string path = RequireLocalPath(style);
            head.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("type", "text/css"),
                new XAttribute("href", RelativeHref(containerPath, path))));
        }
        var document = new XDocument(new XElement(Html + "html", new XAttribute(XNamespace.Xml + "lang", Language), head, body));
        ValidateContent(document, "application/xhtml+xml");
        byte[] chapterBytes = SerializeXml(document, _maximumEntryBytes);
        string navPath = NavigationPath();
        byte[] navBytes = PrepareAppendedNavigation(new EpubNavigationEntry(title, EncodePath(containerPath)), containerPath);
        if (_encryption.Any(encryption => encryption.Path == navPath)) throw new NotSupportedException("Encrypted navigation cannot be edited.");
        XElement manifestItem = PrepareResourceDeclaration(id, containerPath, "application/xhtml+xml", chapterBytes, null);
        var position = new XElement(Opf + "itemref", new XAttribute("idref", id));
        position.SetAttributeValue("linear", linear ? null : "no");
        if (navBytes.LongLength > _maximumEntryBytes) throw new InvalidDataException("Navigation exceeds the retained entry-byte limit.");
        long delta = chapterBytes.LongLength + navBytes.LongLength - _entries[navPath].LongLength;
        EditPackageElement(Root, proposed => {
            proposed.Element(Opf + "manifest")!.Add(new XElement(manifestItem));
            proposed.Element(Opf + "spine")!.Add(new XElement(position));
        }, delta);
        _entries.Add(containerPath, chapterBytes);
        _entries[navPath] = navBytes;
        _retainedBytes += delta;
        return RequireManifestItem(id);
    }

    /// <summary>Adds a UTF-8 stylesheet resource.</summary>
    public EpubManifestItem AddStylesheet(string id, string containerPath, string css) =>
        AddResource(id, containerPath, "text/css", new UTF8Encoding(false, true).GetBytes(css ?? throw new ArgumentNullException(nameof(css))));

    /// <summary>Selects a manifest image as the cover, preserving other item properties.</summary>
    public void SetCoverImage(string manifestId) {
        EpubManifestItem selected = RequireManifestItem(manifestId);
        if (!selected.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Cover must be an image resource.");
        if (PackageVersion == "3.0") {
            EditPackageElement(RequireSection("manifest"), proposed => {
                foreach (XElement item in proposed.Elements(Opf + "item")) {
                    IEnumerable<string> tokens = Tokens((string?)item.Attribute("properties")).Where(token => token != "cover-image");
                    if ((string?)item.Attribute("id") == manifestId) tokens = tokens.Concat(new[] { "cover-image" });
                    item.SetAttributeValue("properties", NormalizeProperties(string.Join(" ", tokens)));
                }
            });
        } else {
            XElement metadata = RequireSection("metadata");
            XElement? cover = metadata.Elements(Opf + "meta").FirstOrDefault(element => (string?)element.Attribute("name") == "cover");
            if (cover == null) EditPackageElement(metadata, proposed => proposed.Add(new XElement(Opf + "meta", new XAttribute("name", "cover"), new XAttribute("content", manifestId))));
            else EditPackageElement(cover, proposed => proposed.SetAttributeValue("content", manifestId));
        }
    }

    /// <summary>Adds a distinct reading position for an existing manifest resource.</summary>
    public EpubSpineItem AddSpineItem(string manifestId, bool linear = true, string? properties = null) {
        RequireManifestItem(manifestId);
        VerifyPropertiesVersion(properties);
        if (Spine.Any(position => position.ManifestId == manifestId))
            throw new InvalidOperationException("New spine positions must not repeat a manifest id. Add a separate chapter resource for repeated content.");
        var element = new XElement(Opf + "itemref", new XAttribute("idref", manifestId));
        element.SetAttributeValue("linear", linear ? null : "no"); element.SetAttributeValue("properties", NormalizeProperties(properties));
        EditPackageElement(RequireSection("spine"), proposed => proposed.Add(new XElement(element)));
        return new EpubSpineItem(RequireSection("spine").Elements(Opf + "itemref").Last(), this);
    }
    /// <summary>Moves a reading position while preserving its attributes and repeated references.</summary>
    public void MoveSpineItem(int fromIndex, int toIndex) {
        XElement[] items = RequireSection("spine").Elements(Opf + "itemref").ToArray();
        if (fromIndex < 0 || fromIndex >= items.Length || toIndex < 0 || toIndex >= items.Length) throw new ArgumentOutOfRangeException();
        if (fromIndex == toIndex) return;
        EditPackageElement(RequireSection("spine"), proposed => {
            XElement[] positions = proposed.Elements(Opf + "itemref").ToArray();
            XElement moved = positions[fromIndex]; moved.Remove();
            if (fromIndex < toIndex) positions[toIndex].AddAfterSelf(moved); else positions[toIndex].AddBeforeSelf(moved);
        });
    }
    /// <summary>Removes one reading position without deleting its resource.</summary>
    public void RemoveSpineItem(int index) {
        XElement[] items = RequireSection("spine").Elements(Opf + "itemref").ToArray();
        if (index < 0 || index >= items.Length) throw new ArgumentOutOfRangeException(nameof(index));
        EditPackageElement(RequireSection("spine"), proposed => proposed.Elements(Opf + "itemref").ElementAt(index).Remove());
    }
    /// <summary>Removes an unreferenced declaration and its payload. Callers must first update content/navigation links.</summary>
    public void RemoveResource(string manifestId) {
        EpubManifestItem item = RequireManifestItem(manifestId);
        string path = RequireLocalPath(item);
        EnsureResourceMutationAllowed(path, removing: true);
        if (Spine.Any(position => position.ManifestId == manifestId) || Manifest.Any(resource => resource.FallbackId == manifestId || resource.FallbackStyleId == manifestId || resource.MediaOverlayId == manifestId) ||
            Root.Descendants().Where(element => element.Name == Opf + "meta" || element.Name == Opf + "link")
                .Attributes("refines").Any(attribute => ReferencesPackageId(attribute.Value, manifestId)) ||
            PackageResourceReferences(Root).Any(attribute => EpubReference.Resolve(PackagePath, attribute.Value).ContainerPath == path) ||
            Root.Element(Opf + "bindings")?.Elements(Opf + "mediaType").Any(binding => (string?)binding.Attribute("handler") == manifestId) == true ||
            HasToken(item.Properties, "nav") || HasToken(item.Properties, "cover-image") ||
            RequireSection("metadata").Elements(Opf + "meta").Any(meta => (string?)meta.Attribute("name") == "cover" && (string?)meta.Attribute("content") == manifestId) ||
            HasMediaType(item.MediaType, "application/x-dtbncx+xml") || _encryption.Any(encryption => encryption.Path == path))
            throw new InvalidOperationException("Resource is referenced by package structure or protection metadata.");
        bool removePayload = !Manifest.Any(resource => resource.Id != manifestId && resource.Reference.ContainerPath == path);
        if (removePayload)
            EnsureRemovalPreservesRootfiles(path);
        long releasedBytes = removePayload && _entries.TryGetValue(path, out byte[]? removed) ? removed.LongLength : 0;
        EditPackageElement(RequireSection("manifest"), proposed => proposed.Elements(Opf + "item").Single(element => (string?)element.Attribute("id") == manifestId).Remove(), -releasedBytes);
        if (removePayload) {
            _entries.Remove(path); _retainedBytes -= releasedBytes;
        }
    }

    private EpubManifestItem RequireManifestItem(string id) => Manifest.SingleOrDefault(item => item.Id == id)
        ?? throw new ArgumentException("Manifest id not found: " + id);
    private static string RequireLocalPath(EpubManifestItem item) => item.Reference.Kind == EpubReferenceKind.Container && item.Reference.ContainerPath != null
        ? item.Reference.ContainerPath : throw new NotSupportedException("Remote resources have no retained package payload.");
    private void EnsurePayloadBudget(byte[] data, long replacedLength) {
        if (data.LongLength > _maximumEntryBytes)
            throw new InvalidDataException("Resource exceeds the publication's retained entry/expanded-byte limits.");
        EnsurePackageBudget(_package, data.LongLength - replacedLength);
    }
    private string VerifyContentPath(string path) {
        if (!EpubReader.TryNormalizeArchiveEntryPath(path, out string normalized) || normalized != path || path.EndsWith("/", StringComparison.Ordinal) ||
            path == "mimetype" || path == PackagePath || path.StartsWith("META-INF/", StringComparison.Ordinal)) throw new ArgumentException("Resource requires a canonical, non-reserved container path.", nameof(path));
        var utf8 = new UTF8Encoding(false, true);
        if (utf8.GetByteCount(path) > 65535 || path.Split('/').Any(segment => utf8.GetByteCount(segment) > 255 || segment.EndsWith(".", StringComparison.Ordinal)))
            throw new ArgumentException("Path exceeds OCF component limits or ends with a dot.", nameof(path));
        for (int index = 0; index < path.Length; index++) {
            int code = char.ConvertToUtf32(path, index);
            if (code > 0xffff) index++;
            if (code < 32 || (code >= 127 && code <= 159) || (code >= 0xe000 && code <= 0xf8ff) || code >= 0xf0000 ||
                (code >= 0xfdd0 && code <= 0xfdef) || (code >= 0xfff0 && code <= 0xffff) || (code & 0xffff) >= 0xfffe ||
                (code <= 0xffff && "\"*:<>?\\|".IndexOf((char)code) >= 0)) throw new ArgumentException("Path contains a character forbidden by OCF.", nameof(path));
        }
        return normalized;
    }
    private static string[] Tokens(string? value) => (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
    private static bool HasToken(string? value, string token) => Tokens(value).Contains(token, StringComparer.Ordinal);
    private static bool HasMediaType(string? value, string expected) => string.Equals(value, expected, StringComparison.OrdinalIgnoreCase);
    private static bool IsScriptMediaType(string mediaType) => new[] { "application/javascript", "text/javascript", "application/ecmascript", "text/ecmascript", "text/jscript" }
        .Contains(mediaType.Split(';')[0].Trim(), StringComparer.OrdinalIgnoreCase);
    private static string EncodePath(string path) => string.Join("/", path.Split('/').Select(Uri.EscapeDataString));
    private static string RelativeHref(string ownerPath, string targetPath) =>
        new Uri("epub://package/" + EncodePath(ownerPath)).MakeRelativeUri(new Uri("epub://package/" + EncodePath(targetPath))).OriginalString;
    private static void ValidateContent(XDocument document, string mediaType) {
        XName expected = HasMediaType(mediaType, "application/xhtml+xml") ? Html + "html" :
            HasMediaType(mediaType, "image/svg+xml") ? XName.Get("svg", "http://www.w3.org/2000/svg") : throw new NotSupportedException("Expected XHTML or SVG content.");
        if (document.Root?.Name != expected) throw new InvalidDataException("Content document root does not match its declared media type.");
        if (document.Descendants().Any(element => element.Name.LocalName == "script" ||
            element.Attributes().Any(attribute => attribute.Name.NamespaceName.Length == 0 &&
                (attribute.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase) || attribute.Name.LocalName == "srcdoc"))))
            throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
    }
}
