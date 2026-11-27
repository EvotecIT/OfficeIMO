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

    private readonly OfficeZipPackageReader? _zip;
    private readonly Func<string, OfficePooledBytes>? _readMember;
    private readonly Func<string, long?>? _getMemberLength;
    private readonly bool _validateMetadata;
    // Keyed by part identity, like the members, so names that differ only in non-ASCII case stay separate.
    private readonly Dictionary<string, RelationshipSet> _relationships = new(OfficeAsciiIgnoreCaseComparer.Instance);
    private Dictionary<string, string>? _defaultContentTypes;
    private Dictionary<string, string>? _overrideContentTypes;

    // Part identity (percent-decoded once, ASCII case-insensitive) to the ZIP member that stores the part.
    private readonly Dictionary<string, string> _members;

    private OfficeOpcPackage(IEnumerable<string> names, int count, OfficeZipPackageReader? zip,
        Func<string, OfficePooledBytes>? readMember = null, Func<string, long?>? getMemberLength = null, bool validateMetadata = false) {
        _zip = zip;
        _readMember = readMember;
        _getMemberLength = getMemberLength;
        _validateMetadata = validateMetadata;
        _members = new Dictionary<string, string>(count, OfficeAsciiIgnoreCaseComparer.Instance);
        foreach (string member in names) {
            string identity = ToIdentity(member) ?? member;
            // Two members that name the same part make every reference to it ambiguous.
            if (_members.ContainsKey(identity)) throw new InvalidDataException("The package stores the part '/" + identity + "' more than once.");
            _members.Add(identity, member);
        }
        // Every part needs a content type, so a package without a readable catalog is not a package, whatever is read.
        ReadContentTypes();
    }

    /// <summary>Namespace ids shared by every part read from this package.</summary>
    internal OfficeXmlNamespaces.Table Namespaces { get; } = new();

    /// <exception cref="InvalidDataException">The ZIP container or the content-types part is missing or malformed.</exception>
    internal static OfficeOpcPackage Open(Stream source, bool leaveOpen = false) {
        OfficeZipPackageReader zip = OfficeZipPackageReader.Open(source, leaveOpen);
        try {
            return new OfficeOpcPackage(zip.Names, zip.Count, zip);
        } catch {
            zip.Dispose();
            throw;
        }
    }

    /// <summary>Uses an existing, validated member index. The caller owns the member source.</summary>
    internal static OfficeOpcPackage OpenFromParts(IEnumerable<string> names, int count,
        Func<string, OfficePooledBytes> readMember, Func<string, long?> getMemberLength, bool validateMetadata) =>
        new(names, count, zip: null, readMember, getMemberLength, validateMetadata);

    private bool TryReadMember(string member, out OfficePooledBytes content) {
        if (_zip != null) return _zip.TryRead(member, out content);
        content = _readMember!(member);
        return true;
    }

    private bool TryGetMemberLength(string member, out long length) {
        if (_zip != null) return _zip.TryGetLength(member, out length);
        long? result = _getMemberLength!(member);
        length = result.GetValueOrDefault();
        return result.HasValue;
    }

    /// <summary>Returns whether the package has a part with this name, for example <c>/word/document.xml</c>.</summary>
    internal bool PartExists(string partName) => FindMember(partName) != null;

    /// <summary>Inflates a part into a buffer the caller disposes.</summary>
    /// <exception cref="InvalidDataException">The part is missing or its ZIP member is corrupt.</exception>
    internal OfficePooledBytes ReadPart(string partName) {
        string? member = FindMember(partName);
        if (member == null || !TryReadMember(member, out OfficePooledBytes content)) {
            throw new InvalidDataException("The package has no part '" + partName + "'.");
        }
        return content;
    }

    /// <summary>The uncompressed length a part's ZIP member declares, checked against the content when it is read.</summary>
    /// <exception cref="InvalidDataException">The part is missing.</exception>
    internal long GetPartLength(string partName) {
        string? member = FindMember(partName);
        if (member == null || !TryGetMemberLength(member, out long length)) throw new InvalidDataException("The package has no part '" + partName + "'.");
        return length;
    }

    /// <summary>Whether any XML part declares an uncompressed length above <paramref name="length"/> bytes.</summary>
    internal bool HasXmlPartLongerThan(long length) {
        foreach (KeyValuePair<string, string> entry in _members) {
            if (!TryGetMemberLength(entry.Value, out long partLength) || partLength <= length) continue;
            string? contentType = GetContentType("/" + entry.Value);
            if (contentType != null && contentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    /// <summary>Opens a pull reader over a part. Dispose the returned buffer after reading.</summary>
    internal OfficeXmlPullReader OpenXml(string partName, out OfficePooledBytes content, bool processMarkupCompatibility = true) {
        content = ReadPart(partName);
        return new OfficeXmlPullReader(content.Array, content.Length, Namespaces, processMarkupCompatibility);
    }

    /// <summary>Relationships whose source is <paramref name="sourcePartName"/>, or the package when it is null.</summary>
    internal IReadOnlyList<OfficeOpcRelationship> GetRelationships(string? sourcePartName) => GetRelationshipSet(sourcePartName).All;

    /// <summary>The relationship of <paramref name="sourcePartName"/> (or the package) with this id, or null.</summary>
    internal OfficeOpcRelationship? FindRelationship(string? sourcePartName, string relationshipId) =>
        GetRelationshipSet(sourcePartName).ById.TryGetValue(relationshipId, out OfficeOpcRelationship? relationship) ? relationship : null;

    /// <summary>Checks that a part, usually the main document part, has one of the content types a reader accepts.</summary>
    /// <exception cref="InvalidDataException">The part's content type is missing or not one of <paramref name="contentTypes"/>.</exception>
    internal void RequireContentType(string partName, params string[] contentTypes) {
        string? contentType = GetContentType(partName);
        if (contentType != null) {
            foreach (string accepted in contentTypes) {
                if (string.Equals(contentType, accepted, StringComparison.OrdinalIgnoreCase)) return;
            }
        }
        throw new InvalidDataException("The part '" + partName + "' has the content type '" + (contentType ?? "(none)") + "', which this reader does not accept.");
    }

    private RelationshipSet GetRelationshipSet(string? sourcePartName) {
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
        string identity = GetPartIdentity(relationshipsPart);
        if (_relationships.TryGetValue(identity, out RelationshipSet? cached)) return cached;
        RelationshipSet relationships = PartExists(relationshipsPart) ? ReadRelationships(relationshipsPart) : RelationshipSet.Empty;
        _relationships[identity] = relationships;
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
    /// <summary>Whether the content-types catalog assigns this content type to any part, by override or default.</summary>
    internal bool DeclaresContentType(string contentType) {
        foreach (string declared in _overrideContentTypes!.Values) {
            if (string.Equals(declared, contentType, StringComparison.OrdinalIgnoreCase)) return true;
        }
        foreach (string declared in _defaultContentTypes!.Values) {
            if (string.Equals(declared, contentType, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    internal string? GetContentType(string partName) {
        // Overrides and references may spell a part with different percent-encoding; both compare by identity.
        partName = GetPartIdentity(partName);
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

    /// <summary>The once-decoded spelling two part names share when they name the same part; compare it ASCII
    /// case-insensitively with <see cref="OfficeAsciiIgnoreCaseComparer"/>.</summary>
    internal static string GetPartIdentity(string partName) =>
        partName.Length > 1 && partName[0] == '/' && ToIdentity(partName.Substring(1)) is { } identity ? "/" + identity : partName;

    // Part names are URIs, and a writer may store a member name with or without percent-encoding, so members and
    // references both compare by their once-decoded spelling.
    private static string? ToIdentity(string name) {
        if (name.IndexOf('%') < 0) return name;
        string? decoded = OfficeOpcPartNames.DecodePartIdentity("/" + name);
        return decoded?.Substring(1);
    }

    private RelationshipSet ReadRelationships(string relationshipsPart) {
        var relationships = new List<OfficeOpcRelationship>();
        var ids = new Dictionary<string, OfficeOpcRelationship>(StringComparer.Ordinal);
        string? sourceDirectory = OfficeOpcPartNames.GetRelationshipSourceDirectory(relationshipsPart);
        OfficeXmlPullReader reader = OpenXml(relationshipsPart, out OfficePooledBytes content, processMarkupCompatibility: !_validateMetadata);
        using (content) {
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (_validateMetadata && reader.Depth == 1
                    && !reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationships")) {
                    throw new System.Xml.XmlException("The package relationships part has no Relationships root element.");
                }
                if (reader.Depth != 2 || !reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationship")) continue;
                if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "Id", out string id)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "Type", out string type)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "Target", out string target)) {
                    throw new InvalidDataException("A relationship in '" + relationshipsPart + "' lacks its Id, Type or Target.");
                }
                bool hasMode = reader.TryGetAttribute(OfficeXmlNamespaces.None, "TargetMode", out string mode);
                if (_validateMetadata && hasMode && mode != "Internal" && mode != "External") {
                    throw new System.Xml.XmlException("The package relationship has an invalid TargetMode.",
                        new ArgumentException("TargetMode must be Internal or External."));
                }
                bool external = hasMode && mode == "External";
                if (ids.ContainsKey(id)) throw new InvalidDataException("'" + relationshipsPart + "' repeats the relationship id '" + id + "'.");
                var relationship = new OfficeOpcRelationship(sourceDirectory, id, OfficeXmlNamespaces.NormalizeRelationshipType(type), target, external);
                ids.Add(id, relationship);
                relationships.Add(relationship);
            }
        }
        return new RelationshipSet(relationships.ToArray(), ids);
    }

    private sealed class RelationshipSet {
        internal static readonly RelationshipSet Empty = new(Array.Empty<OfficeOpcRelationship>(), new Dictionary<string, OfficeOpcRelationship>(StringComparer.Ordinal));

        internal RelationshipSet(OfficeOpcRelationship[] all, Dictionary<string, OfficeOpcRelationship> byId) {
            All = all;
            ById = byId;
        }

        internal OfficeOpcRelationship[] All { get; }
        internal Dictionary<string, OfficeOpcRelationship> ById { get; }
    }

    private void ReadContentTypes() {
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
        string? member = FindMember("/" + ContentTypesMember);
        if (member == null || !TryReadMember(member, out OfficePooledBytes content)) {
            throw new InvalidDataException("The package has no content-types part.");
        }
        using (content) {
            var reader = new OfficeXmlPullReader(content.Array, content.Length, Namespaces, processMarkupCompatibility: !_validateMetadata);
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 1
                    && !(reader.NamespaceId == OfficeXmlNamespaces.ContentTypes && reader.LocalNameEquals("Types"))) {
                    throw new InvalidDataException("The content-types part has no Types root element.");
                }
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2
                    || reader.NamespaceId != OfficeXmlNamespaces.ContentTypes) continue;
                bool isDefault = reader.LocalNameEquals("Default");
                if (!isDefault && !reader.LocalNameEquals("Override")) continue;
                if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "ContentType", out string contentType)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, isDefault ? "Extension" : "PartName", out string key)) {
                    if (_validateMetadata) throw new System.Xml.XmlException("The package contains an incomplete content-type declaration.");
                    continue;
                }
                if (isDefault) {
                    if (_validateMetadata) ValidateContentTypeEntry(defaults, key, contentType);
                    defaults[key] = contentType;
                } else {
                    string identity = GetPartIdentity(key);
                    if (_validateMetadata) {
                        if (key.Length < 2 || key[0] != '/' || key.IndexOf("//", StringComparison.Ordinal) >= 0
                            || key.IndexOf('\\') >= 0) throw new ArgumentException("The content-type override has an invalid part name.");
                        ValidateContentTypeEntry(overrides, identity, contentType);
                    }
                    overrides[identity] = contentType;
                }
            }
        }
        _defaultContentTypes = defaults;
        _overrideContentTypes = overrides;
    }

    private static void ValidateContentTypeEntry(Dictionary<string, string> entries, string key, string value) {
        if (entries.ContainsKey(key) || !System.Net.Http.Headers.MediaTypeHeaderValue.TryParse(value, out _)) {
            throw new ArgumentException("The package contains an invalid or duplicate content-type declaration.");
        }
    }

    public void Dispose() => _zip?.Dispose();
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
