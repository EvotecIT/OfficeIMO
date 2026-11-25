using System.IO.Compression;
#if NET8_0_OR_GREATER
using OfficeIMO.SharedSource.Packaging;
#endif

namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// An Open XML package open for editing. Parts stay as their stored ZIP members until something reads them as XML;
/// on save, members nobody touched are copied with their compressed bytes, parsed parts that did not change are
/// written back byte for byte, and only edited markup is serialised again.
/// </summary>
/// <remarks>Part names are absolute (<c>/ppt/slides/slide1.xml</c>) and compare by OPC identity. The
/// content-types catalog and relationship parts are ordinary XML parts here; the helpers below edit them as trees.</remarks>
internal sealed class OfficeOpcEditablePackage : IDisposable {
    internal const string ContentTypesPartName = "/[Content_Types].xml";
    internal const string RelationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";

    private readonly OfficeZipPackageReader? _zip;
    private readonly List<Part> _parts = new();
    private readonly Dictionary<string, Part> _byIdentity = new(OfficeAsciiIgnoreCaseComparer.Instance);
    private readonly DateTime _modified;
    private OfficeXmlTree? _contentTypes;
    private bool _poolTrees;
    // Part identity to its Override element, built on first use.
    private Dictionary<string, int>? _overrides;
    // The catalog's change count when the override index last matched it; any other edit rebuilds the index.
    private int _overridesVersion;
    // Relationships part identity to the ids it uses, built on first use.
    private readonly Dictionary<string, RelationshipIds> _relationshipIds = new(OfficeAsciiIgnoreCaseComparer.Instance);

    private OfficeOpcEditablePackage(OfficeZipPackageReader? zip, DateTime modified) {
        _zip = zip;
        _modified = modified;
    }

    /// <summary>Namespace ids shared by every part of the package.</summary>
    internal OfficeXmlNamespaces.Table Namespaces { get; } = new();

    /// <summary>Opens a package from a readable, seekable stream. The stream must stay open until the package is disposed.</summary>
    /// <param name="source">The package.</param>
    /// <param name="leaveOpen">Whether disposing the package leaves <paramref name="source"/> open.</param>
    /// <param name="poolTrees">Parse parts into storage rented from the array pool, returned when the package is
    /// disposed. Trees from such a package must not be used after it is disposed.</param>
    internal static OfficeOpcEditablePackage Open(Stream source, bool leaveOpen = false, bool poolTrees = false) {
        OfficeZipPackageReader zip = OfficeZipPackageReader.Open(source, leaveOpen);
        try {
            var package = new OfficeOpcEditablePackage(zip, DateTime.Now) { _poolTrees = poolTrees };
            for (int index = 0; index < zip.Count; index++) {
                string member = zip.GetName(index);
                string identity = Identity("/" + member);
                if (package._byIdentity.ContainsKey(identity)) throw new InvalidDataException("The package stores the part '" + identity + "' more than once.");
                // The part keeps the member's spelling; only the lookup key is decoded.
                package.Add(new Part("/" + member, member) { FromSource = true });
            }
            return package;
        } catch {
            zip.Dispose();
            throw;
        }
    }

    /// <summary>Creates an empty package with a content-types catalog that knows XML and relationship parts.</summary>
    internal static OfficeOpcEditablePackage Create(DateTime? modified = null) {
        var package = new OfficeOpcEditablePackage(null, modified ?? DateTime.Now);
        OfficeXmlTree types = OfficeXmlTree.Create(package.Namespaces, OfficeXmlNamespaces.ContentTypes, "Types", 64);
        types.DeclareNamespace(types.Root, string.Empty, "http://schemas.openxmlformats.org/package/2006/content-types");
        package.Add(new Part(ContentTypesPartName, "[Content_Types].xml") { Tree = types });
        package._contentTypes = types;
        package.AddDefaultContentType("rels", RelationshipContentType);
        package.AddDefaultContentType("xml", "application/xml");
        return package;
    }

    internal bool PartExists(string partName) => _byIdentity.TryGetValue(Identity(partName), out Part? part) && !part.Deleted;

    /// <summary>Part names in package order, the content-types catalog included.</summary>
    internal IEnumerable<string> PartNames {
        get {
            foreach (Part part in _parts) {
                if (!part.Deleted) yield return part.Name;
            }
        }
    }

    /// <summary>The XML tree of a part, parsed on first use.</summary>
    /// <exception cref="InvalidDataException">The package has no such part.</exception>
    /// <summary>The part's tree when it has already been parsed or created, without parsing it; otherwise null.</summary>
    internal OfficeXmlTree? GetParsedXml(string partName) => Find(partName)?.Tree;

    /// <summary>
    /// A pull reader over a part the package has not parsed, edited or replaced, reading its stored bytes without building
    /// a tree; false for any other part, which <see cref="GetXml"/> reads. The caller disposes <paramref name="content"/>.
    /// </summary>
    internal bool TryReadUnparsed(string partName, out OfficeXmlPullReader reader, out OfficePooledBytes content) {
        reader = default!;
        content = default;
        Part? part = Find(partName);
        if (part == null || part.Tree != null || part.Content != null || !part.FromSource || !_zip!.TryRead(part.Member, out content)) return false;
        try {
            reader = new OfficeXmlPullReader(content.Array, content.Length, Namespaces);
        } catch {
            content.Dispose();
            content = default;
            throw;
        }
        return true;
    }

    internal OfficeXmlTree GetXml(string partName) {
        Part part = Find(partName) ?? throw new InvalidDataException("The package has no part '" + partName + "'.");
        if (part.Tree != null) return part.Tree;
        if (part.Content != null) {
            part.Tree = Parse(part.Content, part.ContentLength);
            return part.Tree;
        }
        if (!_zip!.TryRead(part.Member, out OfficePooledBytes content)) throw new InvalidDataException("The package has no part '" + partName + "'.");
        part.Source = content;
        part.Tree = Parse(content.Array, content.Length);
        return part.Tree;
    }

    private OfficeXmlTree Parse(byte[] data, int length) =>
        _poolTrees ? OfficeXmlTree.ParsePooled(data, length, Namespaces) : OfficeXmlTree.Parse(data, length, Namespaces);

    /// <summary>Adds an XML part with a content-type override and returns its tree.</summary>
    internal OfficeXmlTree CreateXmlPart(string partName, string contentType, int namespaceId, string rootName, int capacity = 32) {
        OfficeXmlTree tree = OfficeXmlTree.Create(Namespaces, namespaceId, rootName, capacity);
        AddPart(partName, contentType).Tree = tree;
        return tree;
    }

    /// <summary>A part's bytes as the package would save them: its tree written out, its content, or its stored bytes.</summary>
    /// <exception cref="InvalidDataException">The package has no such part.</exception>
    internal byte[] ReadPartBytes(string partName) {
        Part part = Find(partName) ?? throw new InvalidDataException("The package has no part '" + partName + "'.");
        using var output = new MemoryStream();
        WritePart(part, output, new OfficeXmlUtf8Writer());
        return output.ToArray();
    }

    /// <summary>Adds a part with binary content, such as an image.</summary>
    internal void CreateBinaryPart(string partName, string? contentType, byte[] content) {
        Part part = AddPart(partName, contentType);
        part.Content = content;
        part.ContentLength = content.Length;
    }

    /// <summary>The bytes of a part as the package holds them now; the caller disposes the result.</summary>
    /// <exception cref="InvalidDataException">The package has no such part.</exception>
    internal OfficePooledBytes ReadPart(string partName) {
        Part part = Find(partName) ?? throw new InvalidDataException("The package has no part '" + partName + "'.");
        if (part.Tree != null) throw new InvalidOperationException("The part '" + partName + "' is held as an XML tree.");
        if (part.Content != null) {
            OfficePooledBytes copy = OfficePooledBytes.Rent(part.ContentLength);
            Buffer.BlockCopy(part.Content, 0, copy.Array, 0, part.ContentLength);
            return copy;
        }
        if (!_zip!.TryRead(part.Member, out OfficePooledBytes content)) throw new InvalidDataException("The package has no part '" + partName + "'.");
        return content;
    }

    /// <summary>Replaces the content of an existing part with <paramref name="length"/> bytes of <paramref name="content"/>.</summary>
    /// <exception cref="InvalidDataException">The package has no such part.</exception>
    internal void ReplacePart(string partName, byte[] content, int length) => ReplacePart(partName, content, length, pooled: false);

    /// <summary>
    /// Replaces the content of an existing part; with <paramref name="pooled"/>, <paramref name="content"/> was rented from
    /// the shared array pool and goes back to it when the package is disposed (on .NET 8 and later; older targets let it go).
    /// </summary>
    internal void ReplacePart(string partName, byte[] content, int length, bool pooled) {
        Part part = Find(partName) ?? throw new InvalidDataException("The package has no part '" + partName + "'.");
        part.Source?.Dispose();
        part.Source = null;
        ReturnContent(part);
        part.Tree = null;
        part.Content = content;
        part.ContentLength = length;
        part.ContentPooled = pooled;
        part.FromSource = false;
    }

    private static void ReturnContent(Part part) {
#if NET8_0_OR_GREATER
        if (part.ContentPooled && part.Content != null) System.Buffers.ArrayPool<byte>.Shared.Return(part.Content);
#endif
        part.ContentPooled = false;
        part.Content = null;
    }

    /// <summary>Removes a part and its content-type override. Relationships that point to it are the caller's to remove.</summary>
    internal void DeletePart(string partName) {
        Part? part = Find(partName);
        if (part == null) return;
        part.DeletedContentType = GetContentType(part.Name);
        part.Deleted = true;
        RemoveOverride(part.Name);
    }

    /// <summary>The stored length of a part the package has not parsed or replaced; false for any other part.</summary>
    internal bool TryGetStoredLength(string partName, out long length) {
        length = 0;
        Part? part = Find(partName);
        return part is { Tree: null, Content: null, FromSource: true } && _zip!.TryGetLength(part.Member, out length);
    }

    /// <summary>Whether a part is still the one stored in the package that was opened, neither replaced nor given new content.</summary>
    internal bool IsStoredPart(string partName) => Find(partName) is { FromSource: true, Content: null };

    /// <summary>A part's content type: its override, or the default for its extension; null when neither is declared.</summary>
    internal string? GetContentType(string partName) {
        if (Overrides().TryGetValue(Identity(partName), out int node)) return ContentTypes.GetAttribute(node, OfficeXmlNamespaces.None, "ContentType");
        int dot = partName.LastIndexOf('.');
        if (dot < 0 || dot < partName.LastIndexOf('/')) return null;
        string extension = partName.Substring(dot + 1);
        OfficeXmlTree types = ContentTypes;
        for (int child = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Default"); child != OfficeXmlTree.NoNode;
            child = types.FindNextSibling(child, OfficeXmlNamespaces.ContentTypes, "Default")) {
            if (string.Equals(types.GetAttribute(child, OfficeXmlNamespaces.None, "Extension"), extension, StringComparison.OrdinalIgnoreCase)) {
                return types.GetAttribute(child, OfficeXmlNamespaces.None, "ContentType");
            }
        }
        return null;
    }

    /// <summary>How many parts have, or had before they were deleted, a content type.</summary>
    internal int CountParts(string contentType) {
        int count = 0;
        foreach (Part part in _parts) {
            string? type = part.Deleted ? part.DeletedContentType : GetContentType(part.Name);
            if (string.Equals(type, contentType, StringComparison.OrdinalIgnoreCase)) count++;
        }
        return count;
    }

    // ---- Content types --------------------------------------------------------------------------------------

    internal OfficeXmlTree ContentTypes => _contentTypes ??= GetXml(ContentTypesPartName);

    internal void AddDefaultContentType(string extension, string contentType) {
        OfficeXmlTree types = ContentTypes;
        for (int node = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Default"); node != OfficeXmlTree.NoNode;
            node = types.FindNextSibling(node, OfficeXmlNamespaces.ContentTypes, "Default")) {
            if (string.Equals(types.GetAttribute(node, OfficeXmlNamespaces.None, "Extension"), extension, StringComparison.OrdinalIgnoreCase)) return;
        }
        int added = types.CreateElement(OfficeXmlNamespaces.ContentTypes, "Default");
        types.SetAttribute(added, OfficeXmlNamespaces.None, "Extension", extension);
        types.SetAttribute(added, OfficeXmlNamespaces.None, "ContentType", contentType);
        // Defaults come before overrides, as Office writes them.
        int firstOverride = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Override");
        types.InsertBefore(types.Root, added, firstOverride);
    }

    internal void SetOverride(string partName, string contentType) {
        OfficeXmlTree types = ContentTypes;
        Dictionary<string, int> overrides = Overrides();
        string identity = Identity(partName);
        if (!overrides.TryGetValue(identity, out int node)) {
            node = types.CreateElement(OfficeXmlNamespaces.ContentTypes, "Override");
            types.SetAttribute(node, OfficeXmlNamespaces.None, "PartName", partName);
            types.Append(types.Root, node);
            overrides.Add(identity, node);
        }
        types.SetAttribute(node, OfficeXmlNamespaces.None, "ContentType", contentType);
        _overridesVersion = types.Version;
    }

    private void RemoveOverride(string partName) {
        Dictionary<string, int> overrides = Overrides();
        string identity = Identity(partName);
        if (!overrides.TryGetValue(identity, out int node)) return;
        ContentTypes.Remove(node);
        overrides.Remove(identity);
        _overridesVersion = ContentTypes.Version;
    }

    // Callers can edit the catalog through ContentTypes, so the index is trusted only while the tree is as it left it.
    private Dictionary<string, int> Overrides() {
        OfficeXmlTree types = ContentTypes;
        if (_overrides != null && _overridesVersion == types.Version) return _overrides;
        _overridesVersion = types.Version;
        var overrides = new Dictionary<string, int>(OfficeAsciiIgnoreCaseComparer.Instance);
        for (int node = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Override"); node != OfficeXmlTree.NoNode;
            node = types.FindNextSibling(node, OfficeXmlNamespaces.ContentTypes, "Override")) {
            string? name = types.GetAttribute(node, OfficeXmlNamespaces.None, "PartName");
            // The first Override for a part is the one readers apply.
            if (name != null && !overrides.ContainsKey(Identity(name))) overrides.Add(Identity(name), node);
        }
        return _overrides = overrides;
    }

    // ---- Relationships --------------------------------------------------------------------------------------

    /// <summary>The relationships part of a source part (null for the package), created when it is missing.</summary>
    internal OfficeXmlTree GetRelationships(string? sourcePartName) => RelationshipsTree(OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName));

    private OfficeXmlTree RelationshipsTree(string name) {
        if (Find(name) != null) return GetXml(name);
        OfficeXmlTree tree = OfficeXmlTree.Create(Namespaces, OfficeXmlNamespaces.PackageRelationships, "Relationships", 4);
        tree.DeclareNamespace(tree.Root, string.Empty, "http://schemas.openxmlformats.org/package/2006/relationships");
        // A package may type its relationship parts with overrides alone, so the new part needs the default.
        AddDefaultContentType("rels", RelationshipContentType);
        AddPart(name, null).Tree = tree;
        return tree;
    }

    /// <summary>Adds a relationship and returns its id: <paramref name="id"/>, or the first free <c>rIdN</c>.</summary>
    internal string AddRelationship(string? sourcePartName, string type, string target, bool external = false, string? id = null) {
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
        OfficeXmlTree tree = RelationshipsTree(relationshipsPart);
        RelationshipIds ids = RelationshipIdsOf(relationshipsPart, tree);
        if (id != null && ids.Used.Contains(id)) throw new InvalidOperationException("'" + relationshipsPart + "' already has the relationship id '" + id + "'.");
        if (id == null) {
            do {
                id = "rId" + ids.Next.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ids.Next++;
            } while (ids.Used.Contains(id));
        }
        ids.Used.Add(id);
        int relationship = tree.CreateElement(OfficeXmlNamespaces.PackageRelationships, "Relationship");
        tree.SetAttribute(relationship, OfficeXmlNamespaces.None, "Id", id);
        tree.SetAttribute(relationship, OfficeXmlNamespaces.None, "Type", type);
        tree.SetAttribute(relationship, OfficeXmlNamespaces.None, "Target", target);
        if (external) tree.SetAttribute(relationship, OfficeXmlNamespaces.None, "TargetMode", "External");
        tree.Append(tree.Root, relationship);
        ids.Version = tree.Version;
        return id;
    }

    /// <summary>The target part of a relationship by id, resolved against the source part, or null.</summary>
    internal string? GetRelationshipTarget(string? sourcePartName, string id) {
        string name = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
        if (Find(name) == null) return null;
        OfficeXmlTree tree = GetXml(name);
        for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
            node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
            if (!string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id"), id, StringComparison.Ordinal)) continue;
            if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) return null;
            return OfficeOpcPartNames.ResolveRelationshipTarget(name, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
        }
        return null;
    }

    private RelationshipIds RelationshipIdsOf(string relationshipsPart, OfficeXmlTree tree) {
        string identity = Identity(relationshipsPart);
        // Callers can edit the relationships through their tree, so cached ids hold only while it is as they left it.
        if (_relationshipIds.TryGetValue(identity, out RelationshipIds? ids) && ids.Tree == tree && ids.Version == tree.Version) return ids;
        ids = new RelationshipIds { Tree = tree, Version = tree.Version };
        for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
            node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
            string? existing = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
            if (existing != null) ids.Used.Add(existing);
        }
        ids.Next = ids.Used.Count + 1;
        _relationshipIds[identity] = ids;
        return ids;
    }

    private sealed class RelationshipIds {
        internal HashSet<string> Used { get; } = new(StringComparer.Ordinal);
        internal int Next { get; set; }
        internal OfficeXmlTree? Tree { get; set; }
        internal int Version { get; set; }
    }

    // Part names without percent-encoding are already their identity, which spares a copy per lookup.
    private static string Identity(string partName) => partName.IndexOf('%') < 0 ? partName : OfficeOpcPackage.GetPartIdentity(partName);

    // ---- Save -----------------------------------------------------------------------------------------------

    /// <summary>Writes the package to <paramref name="destination"/>, which is not disposed.</summary>
    internal void Save(Stream destination) {
        ArgumentNullException.ThrowIfNull(destination);
        var xml = new OfficeXmlUtf8Writer();
#if NET8_0_OR_GREATER
        if (BitConverter.IsLittleEndian) {
            using var zip = new OfficeZipWriter(destination, _modified);
            // Large replaced parts are deflated side by side first, then appended in package order.
            Dictionary<Part, OfficeZipWriter.PreparedEntry>? prepared = PrepareLargeParts();
            foreach (Part part in _parts) {
                if (part.Deleted) continue;
                if (prepared != null && prepared.TryGetValue(part, out OfficeZipWriter.PreparedEntry? preparedEntry)) {
                    zip.WritePrepared(part.Member, preparedEntry);
                    continue;
                }
                bool unchanged = part.FromSource && (part.Tree == null || !part.Tree.IsChanged);
                if (unchanged && _zip!.TryReadStored(part.Member, out OfficeZipStoredMember stored)) {
                    using (stored) {
                        zip.WriteStored(part.Member, stored.Method, stored.Flags, stored.DosTime, stored.DosDate, stored.Crc, stored.Size,
                            stored.Compressed.Array, 0, stored.Compressed.Length);
                    }
                    continue;
                }
                using Stream entry = zip.OpenEntry(part.Member, CompressionLevel.Fastest);
                WritePart(part, entry, xml);
            }
            return;
        }
#endif
        using var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
        foreach (Part part in _parts) {
            if (part.Deleted) continue;
            ZipArchiveEntry entry = archive.CreateEntry(part.Member, CompressionLevel.Fastest);
            using Stream output = entry.Open();
            WritePart(part, output, xml);
        }
    }

#if NET8_0_OR_GREATER
    private const int ParallelPartBytes = 128 * 1024;

    private Dictionary<Part, OfficeZipWriter.PreparedEntry>? PrepareLargeParts() {
        List<Part>? large = null;
        foreach (Part part in _parts) {
            if (!part.Deleted && part.Tree == null && part.Content != null && part.ContentLength >= ParallelPartBytes) (large ??= new List<Part>()).Add(part);
        }
        if (large == null || large.Count < 2) return null;
        var entries = new OfficeZipWriter.PreparedEntry[large.Count];
        System.Threading.Tasks.Parallel.For(0, large.Count, index => {
            Part part = large[index];
            entries[index] = OfficeZipWriter.Prepare(stream => stream.Write(part.Content!, 0, part.ContentLength));
        });
        var prepared = new Dictionary<Part, OfficeZipWriter.PreparedEntry>(large.Count);
        for (int index = 0; index < large.Count; index++) prepared[large[index]] = entries[index];
        return prepared;
    }
#endif

    private void WritePart(Part part, Stream output, OfficeXmlUtf8Writer xml) {
        if (part.Tree != null) {
            xml.Target = output;
            part.Tree.WriteTo(xml);
            xml.Flush();
            xml.Target = null;
        } else if (part.Content != null) {
            output.Write(part.Content, 0, part.ContentLength);
        } else {
            if (!_zip!.TryRead(part.Member, out OfficePooledBytes content)) throw new InvalidDataException("The package has no part '" + part.Name + "'.");
            using (content) output.Write(content.Array, 0, content.Length);
        }
    }

    public void Dispose() {
        foreach (Part part in _parts) {
            if (_poolTrees) part.Tree?.ReleasePooledStorage();
            part.Source?.Dispose();
            part.Source = null;
            ReturnContent(part);
        }
        _zip?.Dispose();
    }

    private Part AddPart(string partName, string? contentType) {
        if (partName.Length < 2 || partName[0] != '/') throw new ArgumentException("A part name starts with '/'.", nameof(partName));
        string identity = Identity(partName);
        if (_byIdentity.TryGetValue(identity, out Part? existing) && !existing.Deleted) throw new InvalidOperationException("The package already has the part '" + partName + "'.");
        var part = new Part(partName, partName.Substring(1));
        if (existing != null) {
            _parts[_parts.IndexOf(existing)] = part;
            _byIdentity[identity] = part;
        } else {
            Add(part);
        }
        if (contentType != null) SetOverride(partName, contentType);
        return part;
    }

    private void Add(Part part) {
        _parts.Add(part);
        _byIdentity[Identity(part.Name)] = part;
    }

    private Part? Find(string partName) =>
        _byIdentity.TryGetValue(Identity(partName), out Part? part) && !part.Deleted ? part : null;

    private sealed class Part {
        internal Part(string name, string member) {
            Name = name;
            Member = member;
        }

        internal string Name { get; }
        internal string Member { get; }
        internal OfficeXmlTree? Tree { get; set; }
        internal OfficePooledBytes? Source { get; set; }
        internal byte[]? Content { get; set; }
        internal int ContentLength { get; set; }
        internal bool Deleted { get; set; }
        internal string? DeletedContentType { get; set; }
        internal bool ContentPooled { get; set; }
        // Stored in the package that was opened, so an unchanged part can be copied as its ZIP member.
        internal bool FromSource { get; set; }
    }
}
