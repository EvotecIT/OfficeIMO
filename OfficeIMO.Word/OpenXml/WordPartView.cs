using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>A relationship of a part: an internal one names its target part, an external one its URI as written.</summary>
internal readonly struct WordPartRelationship {
    internal WordPartRelationship(string id, string type, string? target, bool isExternal) {
        Id = id;
        Type = type;
        Target = target;
        IsExternal = isExternal;
    }

    internal string Id { get; }
    internal string Type { get; }
    internal string? Target { get; }
    internal bool IsExternal { get; }
}

/// <summary>
/// The parts of a Word package as readers that walk its markup see them, over whatever holds the package: OfficeIMO's
/// own package while the document works without the Open XML SDK, or the SDK package once it is open. Each part is
/// read through <see cref="WordPartView"/>, which hands out its elements as <see cref="OfficeXmlNode"/>. Reading never
/// opens the SDK package.
/// </summary>
internal sealed partial class WordPackageView {
    internal const string HeaderRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "header";
    internal const string FooterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footer";
    internal const string FootnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footnotes";
    internal const string EndnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "endnotes";
    internal const string NumberingRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "numbering";
    internal const string StylesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles";
    internal const string StylesWithEffectsRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
    internal const string SettingsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "settings";
    internal const string CommentsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments";
    internal const string ImageRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "image";
    internal const string HyperlinkRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "hyperlink";

    private readonly OfficeOpcEditablePackage? _owned;
    private readonly Dictionary<string, WordPartView> _ownedParts = new(OfficeAsciiIgnoreCaseComparer.Instance);
    private readonly Dictionary<object, WordPartView> _sdkParts = new(OfficeIMO.Internal.OfficeReferenceEqualityComparer<object>.Instance);
    private WordPartView? _main;
    private bool _mainRead;

    private WordPackageView(WordDocument? document, OfficeOpcEditablePackage? owned, object? sdkDocument) {
        Document = document;
        _owned = owned;
        SdkDocument = sdkDocument;
    }

    /// <summary>The document the package belongs to, or null for an SDK package read on its own.</summary>
    internal WordDocument? Document { get; }

    /// <summary>The SDK package, or null while the document works on OfficeIMO's own package.</summary>
    internal object? SdkDocument { get; }

    internal bool IsSdk => SdkDocument != null;

    /// <summary>
    /// The package of <paramref name="document"/> as it is held now. A document that opens the SDK package later needs a
    /// new view; elements read through this one belong to the package it was made over.
    /// </summary>
    internal static WordPackageView For(WordDocument document) {
        if (document._ownedModel is { } model) return new WordPackageView(document, model.Package, null);
        return new WordPackageView(document, null, SdkPackageOf(document));
    }

    /// <summary>An SDK package read on its own, for engine code that holds only the SDK package.</summary>
    internal static WordPackageView ForSdkPackage(object sdkDocument) => new(null, null, sdkDocument);

    /// <summary>The main document part, or null when the package has none.</summary>
    internal WordPartView? Main {
        get {
            if (!_mainRead) {
                _mainRead = true;
                if (_owned != null) {
                    _main = Document?._ownedModel is { } model ? OwnedPart(model.MainPartName) : null;
                } else {
                    object? part = SdkMainPart(SdkDocument!);
                    _main = part == null ? null : SdkPart(part);
                }
            }
            return _main;
        }
    }

    /// <summary>
    /// The view of a part of OfficeIMO's own package. Two spellings of one part name, such as <c>header1.xml</c> and
    /// <c>%68eader1.xml</c>, share one view, as they share one part; the view keeps the spelling it was first asked for.
    /// </summary>
    internal WordPartView OwnedPart(string partName) {
        string identity = partName.IndexOf('%') < 0 ? partName : OfficeOpcPackage.GetPartIdentity(partName);
        if (!_ownedParts.TryGetValue(identity, out WordPartView? part)) {
            part = new WordPartView(this, partName, null);
            _ownedParts.Add(identity, part);
        }
        return part;
    }

    internal WordPartView SdkPart(object sdkPart) {
        if (!_sdkParts.TryGetValue(sdkPart, out WordPartView? part)) {
            part = new WordPartView(this, SdkPartName(sdkPart), sdkPart);
            _sdkParts.Add(sdkPart, part);
        }
        return part;
    }

    internal OfficeOpcEditablePackage OwnedPackage => _owned ?? throw new InvalidOperationException("The package is held by the Open XML SDK.");

    /// <summary>The text of a relationship target URI as the package API presents it.</summary>
    internal static string UriText(string target) =>
        Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri.ToString() : target;
}

/// <summary>One part of a <see cref="WordPackageView"/>: its root element, relationships and content.</summary>
internal sealed class WordPartView {
    private OfficeXmlNode _root;
    private bool _rootRead;
    private List<WordPartRelationship>? _relationships;
    private Dictionary<string, int>? _relationshipIndexes;

    internal WordPartView(WordPackageView package, string partName, object? sdkPart) {
        Package = package;
        PartName = partName;
        SdkPart = sdkPart;
    }

    internal WordPackageView Package { get; }

    /// <summary>The part name, such as <c>/word/document.xml</c>, as the part's URI reads.</summary>
    internal string PartName { get; }

    /// <summary>The SDK part, or null for a part of OfficeIMO's own package.</summary>
    internal object? SdkPart { get; }

    internal string? ContentType => SdkPart != null ? WordPackageView.SdkContentType(SdkPart) : Package.OwnedPackage.GetContentType(PartName);

    /// <summary>The root element of the part, or <see cref="OfficeXmlNode.None"/> for a part without markup.</summary>
    internal OfficeXmlNode Root {
        get {
            if (!_rootRead) {
                _rootRead = true;
                if (SdkPart != null && WordPackageView.SdkIsUnloadedStylesPart(SdkPart)) {
                    // A styles part the SDK hasn't loaded is read from its content, so reading it doesn't load it; it
                    // holds no unsaved edits, and readers never edit it.
                    byte[] content = ReadBytes();
                    OfficeXmlTree tree = OfficeXmlTree.Parse(content, content.Length, new OfficeXmlNamespaces.Table());
                    _root = new OfficeXmlNode(tree, tree.Root);
                } else if (SdkPart != null) {
                    _root = OfficeXmlNode.FromSdk(WordPackageView.SdkRootElement(SdkPart));
                } else {
                    OfficeXmlTree tree = Package.OwnedPackage.GetXml(PartName);
                    _root = tree.Root == OfficeXmlTree.NoNode ? default : new OfficeXmlNode(tree, tree.Root);
                }
            }
            return _root;
        }
    }

    /// <summary>The root element when it has this name, as a typed SDK part root reads.</summary>
    internal OfficeXmlNode RootNamed(int namespaceId, string localName) {
        OfficeXmlNode root = Root;
        return root.Is(namespaceId, localName) ? root : default;
    }

    /// <summary>Every relationship of the part, in the order its relationships part lists them.</summary>
    internal IReadOnlyList<WordPartRelationship> Relationships => _relationships ??= ReadRelationships();

    private List<WordPartRelationship> ReadRelationships() {
        if (SdkPart != null) return WordPackageView.SdkRelationships(SdkPart);
        var relationships = new List<WordPartRelationship>();
        OfficeOpcEditablePackage package = Package.OwnedPackage;
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(PartName);
        if (!package.PartExists(relationshipsPart)) return relationships;
        OfficeXmlTree tree = package.GetXml(relationshipsPart);
        string? sourceDirectory = OfficeOpcPartNames.GetRelationshipSourceDirectory(relationshipsPart);
        for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
            node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
            string? id = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
            string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
            string? target = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target");
            if (id == null || type == null) continue;
            bool external = string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            if (!external && !string.Equals(type, WordPackageView.HyperlinkRelationshipType, StringComparison.Ordinal)) {
                target = OfficeOpcPartNames.ResolveRelationshipTargetFrom(sourceDirectory, target);
            }
            relationships.Add(new WordPartRelationship(id, type, target, external));
        }
        return relationships;
    }

    private bool TryFindRelationship(string id, out WordPartRelationship relationship) {
        IReadOnlyList<WordPartRelationship> relationships = Relationships;
        if (_relationshipIndexes == null) {
            // The first relationship with an id is the one a lookup finds.
            var indexes = new Dictionary<string, int>(relationships.Count, StringComparer.Ordinal);
            for (int index = 0; index < relationships.Count; index++) {
                if (!indexes.ContainsKey(relationships[index].Id)) indexes.Add(relationships[index].Id, index);
            }
            _relationshipIndexes = indexes;
        }
        if (_relationshipIndexes.TryGetValue(id, out int found)) {
            relationship = relationships[found];
            return true;
        }
        relationship = default;
        return false;
    }

    private static bool IsDataPartReference(string type) =>
        type.EndsWith("/audio", StringComparison.Ordinal) || type.EndsWith("/video", StringComparison.Ordinal)
        || type.EndsWith("/media", StringComparison.Ordinal);

    private bool IsPartRelationship(WordPartRelationship relationship) =>
        !relationship.IsExternal && relationship.Target != null
        && !string.Equals(relationship.Type, WordPackageView.HyperlinkRelationshipType, StringComparison.Ordinal)
        && !IsDataPartReference(relationship.Type)
        && (SdkPart != null || Package.OwnedPackage.PartExists(relationship.Target));

    private WordPartView PartOf(WordPartRelationship relationship) =>
        SdkPart != null ? Package.SdkPart(WordPackageView.SdkPartById(SdkPart, relationship.Id)!) : Package.OwnedPart(relationship.Target!);

    /// <summary>The part an internal relationship targets, or null for a missing, external or hyperlink relationship.</summary>
    internal WordPartView? GetPartById(string id) =>
        TryFindRelationship(id, out WordPartRelationship relationship) && IsPartRelationship(relationship) ? PartOf(relationship) : null;

    /// <summary>The relationship type of the part <paramref name="id"/> targets, or null.</summary>
    internal string? PartRelationshipType(string id) =>
        TryFindRelationship(id, out WordPartRelationship relationship) && IsPartRelationship(relationship) ? relationship.Type : null;

    /// <summary>The parts of a relationship type, in relationship order.</summary>
    internal IEnumerable<WordPartView> Parts(string relationshipType) {
        foreach (WordPartRelationship relationship in Relationships) {
            if (string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal) && IsPartRelationship(relationship)) {
                yield return PartOf(relationship);
            }
        }
    }

    /// <summary>Every part this part relates to, in relationship order.</summary>
    internal IEnumerable<WordPartView> ChildParts() {
        foreach (WordPartRelationship relationship in Relationships) {
            if (IsPartRelationship(relationship)) yield return PartOf(relationship);
        }
    }

    /// <summary>The first part of a relationship type, or null.</summary>
    internal WordPartView? FirstPart(string relationshipType) {
        foreach (WordPartView part in Parts(relationshipType)) return part;
        return null;
    }

    /// <summary>A hyperlink relationship's target as its URI reads, and whether it is external; null when there is none.</summary>
    internal string? HyperlinkTarget(string? id, out bool isExternal) {
        isExternal = false;
        if (string.IsNullOrEmpty(id) || !TryFindRelationship(id!, out WordPartRelationship relationship)
            || !string.Equals(relationship.Type, WordPackageView.HyperlinkRelationshipType, StringComparison.Ordinal)) {
            return null;
        }
        isExternal = relationship.IsExternal;
        return WordPackageView.UriText(relationship.Target ?? string.Empty);
    }

    internal string? HyperlinkTarget(string? id) => HyperlinkTarget(id, out _);

    /// <summary>An external relationship that is not a hyperlink: its type and target URI text; false when there is none.</summary>
    internal bool TryGetExternalRelationship(string id, out string type, out string uri) {
        if (TryFindRelationship(id, out WordPartRelationship relationship) && relationship.IsExternal
            && !string.Equals(relationship.Type, WordPackageView.HyperlinkRelationshipType, StringComparison.Ordinal)
            && !IsDataPartReference(relationship.Type)) {
            type = relationship.Type;
            uri = WordPackageView.UriText(relationship.Target ?? string.Empty);
            return true;
        }
        type = string.Empty;
        uri = string.Empty;
        return false;
    }

    internal bool HasExternalRelationship(string id) => TryGetExternalRelationship(id, out _, out _);

    /// <summary>The target of a relationship exactly as its relationships part writes it, or null when there is none.</summary>
    internal string? RawTarget(string id) => TryFindRelationship(id, out WordPartRelationship relationship) ? relationship.Target : null;

    /// <summary>The part's content.</summary>
    internal byte[] ReadBytes() {
        if (SdkPart == null) return Package.OwnedPackage.ReadPartBytes(PartName);
        using Stream stream = WordPackageView.SdkOpenRead(SdkPart);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>The part's content as a stream, without copying the content of an SDK part; the caller disposes it.</summary>
    internal Stream OpenRead() {
        if (SdkPart != null) return WordPackageView.SdkOpenRead(SdkPart);
        OfficeOpcEditablePackage package = Package.OwnedPackage;
        return package.GetParsedXml(PartName) == null
            ? new PooledContentStream(package.ReadPart(PartName))
            : new MemoryStream(package.ReadPartBytes(PartName), writable: false);
    }

    /// <summary>A read-only stream over a part's rented bytes, which it returns when it is disposed.</summary>
    private sealed class PooledContentStream : MemoryStream {
        private OfficePooledBytes _content;

        internal PooledContentStream(OfficePooledBytes content) : base(content.Array, 0, content.Length, writable: false) => _content = content;

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _content.Dispose();
                _content = default;
            }
            base.Dispose(disposing);
        }
    }

    // Writes go to the SDK package only: the redline output is always an SDK document.

    internal string AddImagePart(string contentType, Stream content) => Added(WordPackageView.SdkAddImagePart(RequireSdkPart(), contentType, content));

    internal string AddExternalRelationship(string type, string uri) => Added(WordPackageView.SdkAddExternalRelationship(RequireSdkPart(), type, uri));

    internal string AddHyperlinkRelationship(string uri, bool isExternal) => Added(WordPackageView.SdkAddHyperlinkRelationship(RequireSdkPart(), uri, isExternal));

    // The relationships read so far no longer list everything the part holds.
    private string Added(string id) {
        _relationships = null;
        _relationshipIndexes = null;
        return id;
    }

    private object RequireSdkPart() => SdkPart ?? throw new InvalidOperationException("Parts are added only to an SDK package.");
}

/// <summary>An attribute of an element: its prefix as written, local name, namespace URI and value.</summary>
internal readonly struct WordXmlAttribute {
    internal WordXmlAttribute(string prefix, string localName, string namespaceUri, string value) {
        Prefix = prefix;
        LocalName = localName;
        NamespaceUri = namespaceUri;
        Value = value;
    }

    internal string Prefix { get; }
    internal string LocalName { get; }
    internal string NamespaceUri { get; }
    internal string Value { get; }
}
