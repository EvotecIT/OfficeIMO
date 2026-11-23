namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Read-only Open Packaging Conventions view (ECMA-376 Part 2) over <see cref="OfficeZipPackageReader"/>: part
/// access by part name, the content-types catalog, and relationships resolved to target part names. Strict
/// relationship types are reported with their transitional spelling.
/// </summary>
internal sealed class OfficeOpcPackage : IDisposable {
    internal const string OfficeDocumentRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "officeDocument";
    internal const string SlideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide";

    private const string ContentTypesMember = "[Content_Types].xml";

    private readonly OfficeZipPackageReader _zip;
    private readonly Dictionary<string, OfficeOpcRelationship[]> _relationships = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string>? _defaultContentTypes;
    private Dictionary<string, string>? _overrideContentTypes;

    private OfficeOpcPackage(OfficeZipPackageReader zip) => _zip = zip;

    /// <summary>Namespace ids shared by every part read from this package.</summary>
    internal OfficeXmlNamespaces.Table Namespaces { get; } = new();

    internal static OfficeOpcPackage Open(Stream source, bool leaveOpen = false) =>
        new(OfficeZipPackageReader.Open(source, leaveOpen));

    /// <summary>Returns whether the package has a part with this name, for example <c>/word/document.xml</c>.</summary>
    internal bool PartExists(string partName) => FindMember(partName) != null;

    /// <summary>Inflates a part into a buffer the caller disposes.</summary>
    /// <exception cref="InvalidDataException">The part is missing or its ZIP member is corrupt.</exception>
    internal OfficePooledBytes ReadPart(string partName) {
        string? member = FindMember(partName);
        if (member == null || !_zip.TryRead(member, out OfficePooledBytes content)) {
            throw new InvalidDataException("The package has no part '" + partName + "'.");
        }
        return content;
    }

    /// <summary>Opens a pull reader over a part. Dispose the returned buffer after reading.</summary>
    internal OfficeXmlPullReader OpenXml(string partName, out OfficePooledBytes content, IReadOnlyList<string>? understoodNamespaces = null) {
        content = ReadPart(partName);
        return new OfficeXmlPullReader(content.Array, content.Length, Namespaces, understoodNamespaces);
    }

    /// <summary>Relationships whose source is <paramref name="sourcePartName"/>, or the package when it is null.</summary>
    internal IReadOnlyList<OfficeOpcRelationship> GetRelationships(string? sourcePartName) {
        string relationshipsPart = GetRelationshipsPartName(sourcePartName);
        if (_relationships.TryGetValue(relationshipsPart, out OfficeOpcRelationship[]? cached)) return cached;
        OfficeOpcRelationship[] relationships = PartExists(relationshipsPart)
            ? ReadRelationships(relationshipsPart)
            : Array.Empty<OfficeOpcRelationship>();
        _relationships[relationshipsPart] = relationships;
        return relationships;
    }

    /// <summary>The target part of the first internal relationship of a type, or null when there is none.</summary>
    internal string? GetRelatedPartName(string? sourcePartName, string relationshipType) {
        foreach (OfficeOpcRelationship relationship in GetRelationships(sourcePartName)) {
            if (relationship.TargetPartName != null && string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal)) {
                return relationship.TargetPartName;
            }
        }
        return null;
    }

    /// <summary>The content type of a part from its override or the default for its extension.</summary>
    internal string? GetContentType(string partName) {
        if (_overrideContentTypes == null) ReadContentTypes();
        if (_overrideContentTypes!.TryGetValue(partName, out string? contentType)) return contentType;
        int slash = partName.LastIndexOf('/');
        int dot = partName.LastIndexOf('.');
        if (dot <= slash) return null;
        return _defaultContentTypes!.TryGetValue(partName.Substring(dot + 1), out contentType) ? contentType : null;
    }

    private string? FindMember(string partName) {
        if (partName.Length < 2 || partName[0] != '/') return null;
        string member = partName.Substring(1);
        if (_zip.Contains(member)) return member;
        // Part names are URIs; a writer may have stored the member name with or without percent-encoding.
        if (partName.IndexOf('%') >= 0) {
            string? decoded = OfficePackageSecurityInspector.DecodePartIdentity(partName);
            if (decoded != null && _zip.Contains(decoded.Substring(1))) return decoded.Substring(1);
        }
        return null;
    }

    private static string GetRelationshipsPartName(string? sourcePartName) {
        if (sourcePartName == null || sourcePartName == "/") return "/_rels/.rels";
        int slash = sourcePartName.LastIndexOf('/');
        return sourcePartName.Substring(0, slash + 1) + "_rels/" + sourcePartName.Substring(slash + 1) + ".rels";
    }

    private OfficeOpcRelationship[] ReadRelationships(string relationshipsPart) {
        var relationships = new List<OfficeOpcRelationship>();
        OfficeXmlPullReader reader = OpenXml(relationshipsPart, out OfficePooledBytes content);
        using (content) {
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2
                    || !reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationship")) continue;
                if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "Id", out string id)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "Type", out string type)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "Target", out string target)) {
                    throw new InvalidDataException("A relationship in '" + relationshipsPart + "' lacks its Id, Type or Target.");
                }
                bool external = reader.TryGetAttribute(OfficeXmlNamespaces.None, "TargetMode", out string mode)
                    && string.Equals(mode, "External", StringComparison.Ordinal);
                string? targetPart = external ? null : OfficePackageSecurityInspector.ResolveRelationshipTarget(relationshipsPart, target);
                relationships.Add(new OfficeOpcRelationship(id, OfficeXmlNamespaces.NormalizeRelationshipType(type), target, external, targetPart));
            }
        }
        return relationships.ToArray();
    }

    private void ReadContentTypes() {
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!_zip.TryRead(ContentTypesMember, out OfficePooledBytes content)) {
            throw new InvalidDataException("The package has no content-types part.");
        }
        using (content) {
            var reader = new OfficeXmlPullReader(content.Array, content.Length, Namespaces);
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2
                    || reader.NamespaceId != OfficeXmlNamespaces.ContentTypes) continue;
                if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "ContentType", out string contentType)) continue;
                if (reader.LocalNameEquals("Default") && reader.TryGetAttribute(OfficeXmlNamespaces.None, "Extension", out string extension)) {
                    defaults[extension] = contentType;
                } else if (reader.LocalNameEquals("Override") && reader.TryGetAttribute(OfficeXmlNamespaces.None, "PartName", out string partName)) {
                    overrides[partName] = contentType;
                }
            }
        }
        _defaultContentTypes = defaults;
        _overrideContentTypes = overrides;
    }

    public void Dispose() => _zip.Dispose();
}

/// <summary>One package or part relationship. <see cref="TargetPartName"/> is null for external or unresolvable targets.</summary>
internal sealed class OfficeOpcRelationship {
    internal OfficeOpcRelationship(string id, string type, string target, bool isExternal, string? targetPartName) {
        Id = id;
        Type = type;
        Target = target;
        IsExternal = isExternal;
        TargetPartName = targetPartName;
    }

    internal string Id { get; }
    internal string Type { get; }
    internal string Target { get; }
    internal bool IsExternal { get; }
    internal string? TargetPartName { get; }
}
