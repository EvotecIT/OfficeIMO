namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Read-only Open Packaging Conventions view (ECMA-376 Part 2) over <see cref="OfficeZipPackageReader"/>: part
/// access by part name, the content-types catalog, and relationships resolved to target part names. Strict
/// relationship types are reported with their transitional spelling.
/// </summary>
internal sealed class OfficeOpcPackage : IDisposable {
    internal const string OfficeDocumentRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "officeDocument";
    internal const string SlideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide";
    internal const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    private const string ContentTypesMember = "[Content_Types].xml";

    private readonly OfficeZipPackageReader _zip;
    private readonly Dictionary<string, OfficeOpcRelationship[]> _relationships = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string>? _defaultContentTypes;
    private Dictionary<string, string>? _overrideContentTypes;

    // Part identity (percent-decoded once, ASCII case-insensitive) to the ZIP member that stores the part.
    private readonly Dictionary<string, string> _members;

    private OfficeOpcPackage(OfficeZipPackageReader zip) {
        _zip = zip;
        _members = new Dictionary<string, string>(zip.Count, OfficeAsciiIgnoreCaseComparer.Instance);
        foreach (string member in zip.Names) {
            string identity = ToIdentity(member) ?? member;
            // Two members that name the same part make every reference to it ambiguous.
            if (_members.ContainsKey(identity)) throw new InvalidDataException("The package stores the part '/" + identity + "' more than once.");
            _members.Add(identity, member);
        }
    }

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
    internal OfficeXmlPullReader OpenXml(string partName, out OfficePooledBytes content, bool processMarkupCompatibility = true) {
        content = ReadPart(partName);
        return new OfficeXmlPullReader(content.Array, content.Length, Namespaces, processMarkupCompatibility);
    }

    /// <summary>Relationships whose source is <paramref name="sourcePartName"/>, or the package when it is null.</summary>
    internal IReadOnlyList<OfficeOpcRelationship> GetRelationships(string? sourcePartName) {
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
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

    /// <summary>Title, creator, subject and keywords from the core-properties part; all null when it is absent.</summary>
    internal OfficeOpcCoreProperties ReadCoreProperties() {
        string? partName = GetRelatedPartName(null, CorePropertiesRelationshipType);
        var properties = new OfficeOpcCoreProperties();
        if (partName == null || !PartExists(partName)) return properties;
        OfficeXmlPullReader reader = OpenXml(partName, out OfficePooledBytes content);
        using (content) {
            var text = new OfficeCharBuffer();
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2) continue;
                bool dublinCore = reader.NamespaceId == OfficeXmlNamespaces.DublinCore;
                string? name = dublinCore && reader.LocalNameEquals("title") ? "title"
                    : dublinCore && reader.LocalNameEquals("creator") ? "creator"
                    : dublinCore && reader.LocalNameEquals("subject") ? "subject"
                    : reader.IsElement(OfficeXmlNamespaces.CoreProperties, "keywords") ? "keywords"
                    : null;
                if (name == null) {
                    reader.SkipElement();
                    continue;
                }
                text.Clear();
                int depth = reader.Depth;
                while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
                    if (reader.NodeKind == OfficeXmlNodeKind.Text) reader.AppendText(text);
                }
                string value = text.ToString();
                switch (name) {
                    case "title": properties.Title = value; break;
                    case "creator": properties.Creator = value; break;
                    case "subject": properties.Subject = value; break;
                    default: properties.Keywords = value; break;
                }
            }
        }
        return properties;
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
        string? identity = ToIdentity(partName.Substring(1));
        return identity != null && _members.TryGetValue(identity, out string? member) ? member : null;
    }

    // Part names are URIs, and a writer may store a member name with or without percent-encoding, so members and
    // references both compare by their once-decoded spelling.
    private static string? ToIdentity(string name) {
        if (name.IndexOf('%') < 0) return name;
        string? decoded = OfficeOpcPartNames.DecodePartIdentity("/" + name);
        return decoded?.Substring(1);
    }

    private OfficeOpcRelationship[] ReadRelationships(string relationshipsPart) {
        var relationships = new List<OfficeOpcRelationship>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        string? sourceDirectory = OfficeOpcPartNames.GetRelationshipSourceDirectory(relationshipsPart);
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
                if (!ids.Add(id)) throw new InvalidDataException("'" + relationshipsPart + "' repeats the relationship id '" + id + "'.");
                relationships.Add(new OfficeOpcRelationship(sourceDirectory, id, OfficeXmlNamespaces.NormalizeRelationshipType(type), target, external));
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
    private readonly string? _sourceDirectory;
    private string? _targetPartName;
    private bool _resolved;

    internal OfficeOpcRelationship(string? sourceDirectory, string id, string type, string target, bool isExternal) {
        _sourceDirectory = sourceDirectory;
        Id = id;
        Type = type;
        Target = target;
        IsExternal = isExternal;
    }

    internal string Id { get; }
    internal string Type { get; }
    internal string Target { get; }
    internal bool IsExternal { get; }

    /// <summary>The target part name, resolved on first use: most relationships a reader loads are never followed.</summary>
    internal string? TargetPartName {
        get {
            if (!_resolved) {
                _targetPartName = IsExternal ? null : OfficeOpcPartNames.ResolveRelationshipTargetFrom(_sourceDirectory, Target);
                _resolved = true;
            }
            return _targetPartName;
        }
    }
}

/// <summary>The core document properties a text projection reports.</summary>
internal sealed class OfficeOpcCoreProperties {
    internal string? Title { get; set; }
    internal string? Creator { get; set; }
    internal string? Subject { get; set; }
    internal string? Keywords { get; set; }
}
