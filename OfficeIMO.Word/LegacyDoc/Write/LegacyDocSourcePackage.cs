using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    /// <summary>
    /// The DOCX package the native DOC writer reads: its parts as element trees, parsed on first use, and the
    /// relationships between them. The writer only reads it.
    /// </summary>
    internal sealed class LegacyDocSourcePackage : IDisposable {
        internal const string RelationshipTypePrefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
        internal const string StylesRelationshipType = RelationshipTypePrefix + "styles";
        internal const string NumberingRelationshipType = RelationshipTypePrefix + "numbering";
        internal const string SettingsRelationshipType = RelationshipTypePrefix + "settings";
        internal const string FootnotesRelationshipType = RelationshipTypePrefix + "footnotes";
        internal const string EndnotesRelationshipType = RelationshipTypePrefix + "endnotes";
        internal const string CommentsRelationshipType = RelationshipTypePrefix + "comments";
        internal const string HeaderRelationshipType = RelationshipTypePrefix + "header";
        internal const string FooterRelationshipType = RelationshipTypePrefix + "footer";
        internal const string ThemeRelationshipType = RelationshipTypePrefix + "theme";
        internal const string ImageRelationshipType = RelationshipTypePrefix + "image";
        internal const string HyperlinkRelationshipType = RelationshipTypePrefix + "hyperlink";
        internal const string CommentsExRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
        internal const string CommentsIdsRelationshipType = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
        internal const string PeopleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/people";
        internal const string VbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

        private readonly OfficeOpcEditablePackage _package;
        private readonly bool _ownsPackage;
        private readonly MemoryStream? _source;
        private readonly Dictionary<string, Relationship[]> _relationships = new(StringComparer.OrdinalIgnoreCase);
        private LegacyDocElement _document;
        private LegacyDocStyleCatalog? _styles;

        private LegacyDocSourcePackage(OfficeOpcEditablePackage package, string mainPartName, bool ownsPackage, MemoryStream? source) {
            _package = package;
            MainPartName = mainPartName;
            _ownsPackage = ownsPackage;
            _source = source;
        }

        /// <summary>
        /// Reads the package from DOCX bytes. A positive <paramref name="maxCharactersInPart"/> refuses any part the
        /// writer reads that holds more characters, as the SDK package the bytes came from would.
        /// </summary>
        internal static LegacyDocSourcePackage Open(byte[] bytes, long maxCharactersInPart = 0) {
            var source = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true);
            OfficeOpcEditablePackage? package = null;
            try {
                package = OfficeOpcEditablePackage.Open(source, leaveOpen: true, poolTrees: true);
                var opened = new LegacyDocSourcePackage(package, string.Empty, ownsPackage: true, source) {
                    MaxCharactersInPart = maxCharactersInPart
                };
                string? main = opened.RelatedPart(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
                if (main == null || !package.PartExists(main)) {
                    throw new InvalidOperationException("Document couldn't be saved as legacy DOC because the main document part is missing.");
                }
                opened.MainPartName = main;
                return opened;
            } catch {
                package?.Dispose();
                source.Dispose();
                throw;
            }
        }

        /// <summary>Reads the package a document holds on OfficeIMO's own Open XML layer, without copying it.</summary>
        internal static LegacyDocSourcePackage FromModel(WordPackageModel model) =>
            new(model.Package, model.MainPartName, ownsPackage: false, source: null);

        internal string MainPartName { get; private set; }

        private long MaxCharactersInPart { get; init; }

        /// <summary>The <c>w:document</c> element.</summary>
        internal LegacyDocElement Document {
            get {
                if (_document.IsNone) _document = Root(MainPartName);
                return _document;
            }
        }

        internal LegacyDocElement Body => Document.Child("body");

        /// <summary>The main document part, as the owner of the relationships its content uses.</summary>
        internal LegacyDocPart Main => new(this, MainPartName);

        /// <summary>A part of this package, as the owner of the relationships its content uses.</summary>
        internal LegacyDocPart Part(string partName) => new(this, partName);

        /// <summary>The style definitions of the main part's styles part.</summary>
        internal LegacyDocStyleCatalog Styles => _styles ??= new LegacyDocStyleCatalog(Root(MainPart(StylesRelationshipType)));

        internal LegacyDocElement Settings => Root(MainPart(SettingsRelationshipType));

        internal LegacyDocElement Numbering => Root(MainPart(NumberingRelationshipType));

        internal string? FootnotesPart => MainPart(FootnotesRelationshipType);

        internal string? EndnotesPart => MainPart(EndnotesRelationshipType);

        internal string? CommentsPart => MainPart(CommentsRelationshipType);

        internal LegacyDocElement Footnotes => Root(FootnotesPart);

        internal LegacyDocElement Endnotes => Root(EndnotesPart);

        internal LegacyDocElement Comments => Root(CommentsPart);

        internal LegacyDocElement Theme => Root(MainPart(ThemeRelationshipType));

        /// <summary>The header parts of the main part, once each, in relationship order, with the first id pointing at each.</summary>
        internal IReadOnlyList<LegacyDocRelatedPart> HeaderParts => MainParts(HeaderRelationshipType);

        internal IReadOnlyList<LegacyDocRelatedPart> FooterParts => MainParts(FooterRelationshipType);

        /// <summary>The root element of a part, or none when the part is missing or isn't XML.</summary>
        internal LegacyDocElement Root(string? partName) {
            if (partName == null || !_package.PartExists(partName)) return LegacyDocElement.None;
            if (MaxCharactersInPart > 0 && _package.GetParsedXml(partName) == null) ThrowIfOverCharacterLimit(partName);
            OfficeXmlTree tree = _package.GetXml(partName);
            return new LegacyDocElement(tree, tree.Root);
        }

        private void ThrowIfOverCharacterLimit(string partName) {
            using var content = _package.ReadPart(partName);
            if (content.Length > MaxCharactersInPart && OfficeXmlPullReader.CountCharacters(content.Array, content.Length) > MaxCharactersInPart) {
                throw new InvalidDataException($"The part '{partName}' holds more characters than the package's MaxCharactersInPart limit of {MaxCharactersInPart}.");
            }
        }

        /// <summary>The first part the main part relates to with this type.</summary>
        internal string? MainPart(string relationshipType) => RelatedPart(MainPartName, relationshipType);

        internal IReadOnlyList<LegacyDocRelatedPart> MainParts(string relationshipType) {
            var parts = new List<LegacyDocRelatedPart>();
            foreach (Relationship relationship in Relationships(MainPartName)) {
                if (relationship.IsExternal || relationship.Target == null || !string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal)) continue;
                if (!_package.PartExists(relationship.Target) || parts.Any(part => string.Equals(part.PartName, relationship.Target, StringComparison.OrdinalIgnoreCase))) continue;
                parts.Add(new LegacyDocRelatedPart(relationship.Id, relationship.Target));
            }
            return parts;
        }

        internal string? RelatedPart(string? sourcePartName, string relationshipType) {
            foreach (Relationship relationship in Relationships(sourcePartName)) {
                if (!relationship.IsExternal && relationship.Target != null && string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal)
                    && _package.PartExists(relationship.Target)) {
                    return relationship.Target;
                }
            }
            return null;
        }

        /// <summary>The relationship with this id from a part, or null.</summary>
        internal LegacyDocRelationship? FindRelationship(string sourcePartName, string? id) {
            if (id == null) return null;
            foreach (Relationship relationship in Relationships(sourcePartName)) {
                if (string.Equals(relationship.Id, id, StringComparison.Ordinal)) {
                    return new LegacyDocRelationship(relationship.Id, relationship.Type, relationship.RawTarget, relationship.Target, relationship.IsExternal);
                }
            }
            return null;
        }

        /// <summary>The target part of an internal relationship by id when the part exists, or null.</summary>
        internal string? GetPartById(string sourcePartName, string? id) {
            LegacyDocRelationship? relationship = FindRelationship(sourcePartName, id);
            return relationship is { IsExternal: false, PartName: { } target } && _package.PartExists(target) ? target : null;
        }

        /// <summary>The relationship types of every part the main part reaches, directly or through other parts.</summary>
        internal bool ReachesRelationshipType(Func<string, bool> matches) {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { MainPartName };
            var pending = new Stack<string>();
            pending.Push(MainPartName);
            while (pending.Count > 0) {
                foreach (Relationship relationship in Relationships(pending.Pop())) {
                    if (relationship.IsExternal || relationship.Target == null || !_package.PartExists(relationship.Target)) continue;
                    if (matches(relationship.Type)) return true;
                    if (visited.Add(relationship.Target)) pending.Push(relationship.Target);
                }
            }
            return false;
        }

        /// <summary>The internal parts the main part relates to directly, with their relationship types.</summary>
        internal IEnumerable<(string Type, string PartName)> MainRelatedParts() {
            foreach (Relationship relationship in Relationships(MainPartName)) {
                if (!relationship.IsExternal && relationship.Target != null && _package.PartExists(relationship.Target)) {
                    yield return (relationship.Type, relationship.Target);
                }
            }
        }

        internal string? GetContentType(string partName) => _package.GetContentType(partName);

        internal byte[] ReadPartBytes(string partName) {
            using var content = _package.ReadPart(partName);
            byte[] bytes = new byte[content.Length];
            Buffer.BlockCopy(content.Array, 0, bytes, 0, content.Length);
            return bytes;
        }

        private Relationship[] Relationships(string? sourcePartName) {
            string key = sourcePartName ?? string.Empty;
            if (_relationships.TryGetValue(key, out Relationship[]? cached)) return cached;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
            Relationship[] relationships = Array.Empty<Relationship>();
            if (_package.PartExists(relationshipsPart)) {
                OfficeXmlTree tree = _package.GetXml(relationshipsPart);
                var list = new List<Relationship>();
                for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
                    if (!tree.IsElement(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) continue;
                    string? id = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                    string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                    string? target = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target");
                    if (id == null || type == null) continue;
                    bool external = string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                    string? resolved = external ? null : OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, target);
                    list.Add(new Relationship(id, OfficeXmlNamespaces.NormalizeRelationshipType(type), target, resolved, external));
                }
                relationships = list.ToArray();
            }
            _relationships[key] = relationships;
            return relationships;
        }

        public void Dispose() {
            if (!_ownsPackage) return;
            _package.Dispose();
            _source?.Dispose();
        }

        private sealed class Relationship {
            internal Relationship(string id, string type, string? rawTarget, string? target, bool isExternal) {
                Id = id;
                Type = type;
                RawTarget = rawTarget;
                Target = target;
                IsExternal = isExternal;
            }

            internal string Id { get; }
            internal string Type { get; }
            internal string? RawTarget { get; }
            internal string? Target { get; }
            internal bool IsExternal { get; }
        }
    }

    /// <summary>A part of the source package: the owner its content resolves relationship ids against.</summary>
    internal readonly struct LegacyDocPart {
        internal LegacyDocPart(LegacyDocSourcePackage source, string partName) {
            Source = source;
            PartName = partName;
        }

        internal LegacyDocSourcePackage Source { get; }

        internal string PartName { get; }

        internal bool Exists => Source != null;

        internal LegacyDocElement Root => Source.Root(PartName);

        internal LegacyDocRelationship? FindRelationship(string? id) => Source.FindRelationship(PartName, id);

        /// <summary>The part an internal relationship with this id targets, or null.</summary>
        internal string? GetPartById(string? id) => Source.GetPartById(PartName, id);

        public override string ToString() => PartName;
    }

    internal readonly struct LegacyDocRelatedPart {
        internal LegacyDocRelatedPart(string id, string partName) {
            Id = id;
            PartName = partName;
        }

        internal string Id { get; }

        internal string PartName { get; }
    }

    internal sealed class LegacyDocRelationship {
        internal LegacyDocRelationship(string id, string type, string? target, string? partName, bool isExternal) {
            Id = id;
            Type = type;
            Target = target;
            PartName = partName;
            IsExternal = isExternal;
        }

        internal string Id { get; }

        internal string Type { get; }

        /// <summary>The target as written in the relationships part.</summary>
        internal string? Target { get; }

        /// <summary>The resolved part name of an internal relationship.</summary>
        internal string? PartName { get; }

        internal bool IsExternal { get; }
    }
}
