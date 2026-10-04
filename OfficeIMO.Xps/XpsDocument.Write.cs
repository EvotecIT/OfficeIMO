using OfficeIMO.Core.Internal;

namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    /// <summary>Appends a page to the last fixed document, creating one when the sequence is empty. Dimensions use 1/96-inch units.</summary>
    public XpsPage AddPage(double width = 816, double height = 1056, string language = "en-US") {
        XpsPage.ValidatePageDimension(width); XpsPage.ValidatePageDimension(height);
        if (string.IsNullOrWhiteSpace(language)) throw new ArgumentException("A page language is required.", nameof(language));
        if (_documents.Count > 0) return _documents[_documents.Count - 1].AddPage(width, height, language);
        var document = new XpsFixedDocument(this, NewPartName("Documents/", "/FixedDocument.fdoc"), new XElement(XName.Get("FixedDocument", XpsPackage.Namespace(Format))));
        return document.InsertNewPage(0, width, height, language, attach: true);
    }
    /// <summary>Embeds a caller-provided font and returns its absolute package URI. The caller must have embedding rights.</summary>
    public string AddFont(byte[] fontBytes, bool obfuscate = true) {
        if (fontBytes == null) throw new ArgumentNullException(nameof(fontBytes));
        if (fontBytes.Length < 32) throw new ArgumentException("A font program is required.", nameof(fontBytes));
        string name = "Resources/Fonts/" + Guid.NewGuid().ToString("D") + (obfuscate ? ".odttf" : ".ttf");
        byte[] data = (byte[])fontBytes.Clone();
        if (obfuscate) XpsFontEncoding.Toggle(data, name);
        AddResource(name, data, obfuscate ? "application/vnd.ms-package.obfuscated-opentype" : "application/vnd.ms-opentype");
        return "/" + name;
    }
    /// <summary>Adds an immutable package resource. Existing parts cannot be overwritten through this API.</summary>
    public string AddResource(string partName, byte[] bytes, string contentType) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (string.IsNullOrWhiteSpace(contentType)) throw new ArgumentException("A MIME content type is required.", nameof(contentType));
        string name = XpsPackage.PartName(partName);
        if (name == "[Content_Types].xml" || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) || OfficeOpcPieceAssembler.HasPieceSuffix(name)) throw new ArgumentException("Reserved package part.", nameof(partName));
        if (bytes.Length > _limits.MaximumPartBytes || _parts.Count >= _limits.MaximumParts) throw new InvalidOperationException("XPS resource limit exceeded.");
        if (_parts.ContainsKey(name)) throw new ArgumentException("The resource already exists.", nameof(partName));
        byte[] copy = (byte[])bytes.Clone();
        _ = PrepareOutput(default, new Dictionary<string, byte[]> { [name] = copy }, new Dictionary<string, string> { [name] = contentType }, validateResources: false);
        _parts.Add(name, copy); _types.Add(name, contentType);
        return "/" + name;
    }
    /// <summary>Replaces an existing resource's encoded bytes, retaining its URI and content type. All references see the replacement.</summary>
    /// <remarks>Structural parts and relationships must be edited through their owning APIs. Obfuscated font data must already use the part's native encoding.</remarks>
    public void ReplaceResource(string partName, byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        string name = XpsPackage.PartName(partName);
        if (name == "[Content_Types].xml" || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Use the owning API to edit structural package parts.", nameof(partName));
        _ = Part(name);
        string type = ContentType(name);
        if (type == XpsPackage.Type("fixedpage") || type == XpsPackage.Type("fixeddocument") || type == XpsPackage.Type("fixeddocumentsequence") ||
            type == XpsPackage.Type("documentstructure") || type == XpsPackage.Type("storyfragments"))
            throw new ArgumentException("Use the owning API to edit structural package parts.", nameof(partName));
        if (bytes.Length > _limits.MaximumPartBytes)
            throw new InvalidOperationException("XPS resource limit exceeded.");
        byte[] copy = (byte[])bytes.Clone();
        _ = PrepareOutput(default, new Dictionary<string, byte[]> { [name] = copy }, validateResources: false);
        _parts[name] = copy;
    }
    /// <summary>Returns a copy of a retained package resource.</summary>
    public byte[] GetPartBytes(string partName) => (byte[])Part(XpsPackage.PartName(partName)).Clone();
    /// <summary>Serializes the same dialect. Unknown native parts are retained; signed packages cannot be rewritten.</summary>
    public byte[] Save(CancellationToken cancellationToken = default) {
        using var stream = new MemoryStream(); Save(stream, cancellationToken); return stream.ToArray();
    }
    /// <summary>Atomically replaces a file only after successful serialization.</summary>
    public void Save(string path, CancellationToken cancellationToken = default) => OfficeFileCommit.WriteAtomically(path, s => Save(s, cancellationToken), cancellationToken);
    /// <summary>Writes a ZIP package to a writable stream, leaving it open. Use a fresh/empty stream.</summary>
    public void Save(Stream stream, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        // Rewriting signed XML/ZIP bytes must never imply that signatures remain valid.
        if (_parts.Keys.Any(p => p.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) ||
            _types.Values.Any(t => t.IndexOf("digital-signature", StringComparison.OrdinalIgnoreCase) >= 0))
            throw new NotSupportedException("Saving digitally signed XPS packages is not supported; retain the original signed bytes.");
        var output = PrepareOutput(cancellationToken);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, true);
        foreach (var item in output.OrderBy(p => p.Key, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            var archiveEntry = zip.CreateEntry(item.Key, CompressionLevel.Optimal);
            archiveEntry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using var entry = archiveEntry.Open(); entry.Write(item.Value, 0, item.Value.Length);
        }
    }
    // The same prospective output is used by Save and bounded mutations, including generated OPC metadata.
    private Dictionary<string, byte[]> PrepareOutput(CancellationToken cancellationToken,
        IReadOnlyDictionary<string, byte[]>? replacements = null, IReadOnlyDictionary<string, string>? replacementTypes = null, XpsPage? pending = null, bool validateResources = true) {
        var output = new Dictionary<string, byte[]>(_parts, StringComparer.OrdinalIgnoreCase);
        if (replacements != null) foreach (var replacement in replacements) output[replacement.Key] = replacement.Value;
        IEnumerable<XpsPage> pages = _pageCache.Values;
        if (pending != null) pages = pages.Concat(new[] { pending });
        foreach (XpsPage page in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? markup = null;
            if (replacements != null && replacements.TryGetValue(page.PartName, out var updated)) markup = XpsPackage.Xml(updated, _limits, cancellationToken);
            else output[page.PartName] = page.Serialize();
            string relName = page.PartName.Substring(0, page.PartName.LastIndexOf('/') + 1) + "_rels/" + Path.GetFileName(page.PartName) + ".rels";
            XElement rels = output.TryGetValue(relName, out var relBytes) ? XpsPackage.Xml(relBytes, _limits, cancellationToken) : new XElement(XpsPackage.Relationships + "Relationships");
            foreach (string resource in page.ResourceReferences(markup, name => {
                string resourceType = replacementTypes != null && replacementTypes.TryGetValue(name, out var replacementType) ? replacementType : (_types.TryGetValue(name, out var storedType) ? storedType : "");
                return resourceType == XpsPackage.Type("resourcedictionary") && output.TryGetValue(name, out var dictionaryBytes)
                    ? XpsPackage.Xml(dictionaryBytes, _limits, cancellationToken) : null;
            })) {
                if (validateResources && !output.ContainsKey(resource)) throw new InvalidDataException("Missing XPS part: " + resource);
                string type = XpsPackage.Namespace(Format) + "/required-resource";
                if (!rels.Elements().Any(r => (string?)r.Attribute("Type") == type && XpsPackage.Resolve(page.PartName, (string?)r.Attribute("Target") ?? "") == resource))
                    rels.Add(new XElement(XpsPackage.Relationships + "Relationship", new XAttribute("Id", NextRelationshipId(rels)), new XAttribute("Type", type), new XAttribute("Target", "/" + resource)));
            }
            if (rels.HasElements) output[relName] = XpsPackage.Serialize(rels);
        }
        var types = new XElement(XpsPackage.ContentTypes + "Types", new XElement(XpsPackage.ContentTypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")));
        foreach (string name in output.Keys.Where(n => n != "[Content_Types].xml" && !n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).OrderBy(n => n, StringComparer.Ordinal))
            types.Add(new XElement(XpsPackage.ContentTypes + "Override", new XAttribute("PartName", "/" + name), new XAttribute("ContentType", replacementTypes != null && replacementTypes.TryGetValue(name, out var type) ? type : ContentType(name))));
        output["[Content_Types].xml"] = XpsPackage.Serialize(types);
        if (output.Count > _limits.MaximumParts || output.Values.Any(b => b.Length > _limits.MaximumPartBytes) || output.Values.Sum(b => (long)b.Length) > _limits.MaximumExpandedBytes) throw new InvalidOperationException("XPS output limits exceeded.");
        return output;
    }
    private static string NextRelationshipId(XElement relationships) {
        var ids = new HashSet<string>(relationships.Elements().Select(r => (string?)r.Attribute("Id") ?? ""), StringComparer.Ordinal);
        int index = 1;
        while (ids.Contains("rResource" + index.ToString(CultureInfo.InvariantCulture))) index++;
        return "rResource" + index.ToString(CultureInfo.InvariantCulture);
    }

}

internal static class XpsFontEncoding {
    internal static void Toggle(byte[] bytes, string name) {
        if (bytes.Length < 32 || !Guid.TryParse(Path.GetFileNameWithoutExtension(name), out var guid)) throw new InvalidDataException("Invalid obfuscated XPS font name or payload.");
        // ECMA-388 9.1.7.3: textual GUID octets in reverse order (not Guid.ToByteArray order).
        string hex = guid.ToString("N");
        for (int i = 0; i < 32; i++) bytes[i] ^= byte.Parse(hex.Substring((15 - i % 16) * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
    }
}
