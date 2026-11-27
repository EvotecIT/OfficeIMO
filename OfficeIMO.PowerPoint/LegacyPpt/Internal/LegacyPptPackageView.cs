using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// A copy of a presentation's package read through the Open XML layer: parts by name, their trees, bytes and
    /// relationships. The binary writer and the import fingerprints read the presentation through it, so they see the
    /// same markup whichever model holds the document. Edits made to the copy stay in the copy.
    /// </summary>
    internal sealed class LegacyPptPackageView : IDisposable {
        internal const string RelationshipTypePrefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
        internal const string OfficeDocumentRelationshipType = RelationshipTypePrefix + "officeDocument";
        internal const string SlideRelationshipType = RelationshipTypePrefix + "slide";
        internal const string SlideLayoutRelationshipType = RelationshipTypePrefix + "slideLayout";
        internal const string SlideMasterRelationshipType = RelationshipTypePrefix + "slideMaster";
        internal const string NotesSlideRelationshipType = RelationshipTypePrefix + "notesSlide";
        internal const string NotesMasterRelationshipType = RelationshipTypePrefix + "notesMaster";
        internal const string HandoutMasterRelationshipType = RelationshipTypePrefix + "handoutMaster";
        internal const string ThemeRelationshipType = RelationshipTypePrefix + "theme";
        internal const string ThemeOverrideRelationshipType = RelationshipTypePrefix + "themeOverride";
        internal const string ImageRelationshipType = RelationshipTypePrefix + "image";
        internal const string AudioRelationshipType = RelationshipTypePrefix + "audio";
        internal const string VideoRelationshipType = RelationshipTypePrefix + "video";
        internal const string HyperlinkRelationshipType = RelationshipTypePrefix + "hyperlink";
        internal const string OleObjectRelationshipType = RelationshipTypePrefix + "oleObject";
        internal const string PackageRelationshipType = RelationshipTypePrefix + "package";
        internal const string ChartRelationshipType = RelationshipTypePrefix + "chart";
        internal const string CommentsRelationshipType = RelationshipTypePrefix + "comments";
        internal const string CommentAuthorsRelationshipType = RelationshipTypePrefix + "commentAuthors";
        internal const string ExtendedPropertiesRelationshipType = RelationshipTypePrefix + "extended-properties";
        internal const string CustomPropertiesRelationshipType = RelationshipTypePrefix + "custom-properties";
        internal const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        internal const string VbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        internal const string MediaRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";

        private readonly OfficeOpcEditablePackage _package;
        private readonly Dictionary<string, IReadOnlyList<LegacyPptRelationshipView>> _relationships = new(StringComparer.OrdinalIgnoreCase);

        private LegacyPptPackageView(OfficeOpcEditablePackage package) {
            _package = package;
            Presentation = FindPresentation();
        }

        /// <summary>Copies the document's package as it is now, edits included.</summary>
        internal static LegacyPptPackageView Capture(PowerPointDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            return Open(document.CapturePackageBytes());
        }

        /// <summary>Reads a package from its bytes.</summary>
        internal static LegacyPptPackageView Open(byte[] packageBytes) {
            ArgumentNullException.ThrowIfNull(packageBytes);
            OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(new MemoryStream(packageBytes, writable: false));
            try {
                return new LegacyPptPackageView(package);
            } catch {
                package.Dispose();
                throw;
            }
        }

        /// <summary>The presentation part, or none when the package has no main part.</summary>
        internal LegacyPptPartView Presentation { get; }

        /// <summary>The part with this name, or none.</summary>
        internal LegacyPptPartView Part(string? partName) =>
            partName != null && _package.PartExists(partName) ? new LegacyPptPartView(this, partName) : default;

        /// <summary>The package-level relationships.</summary>
        internal IReadOnlyList<LegacyPptRelationshipView> PackageRelationships => RelationshipsOf(null);

        /// <summary>The part a package-level relationship of this type targets, or none.</summary>
        internal LegacyPptPartView PackageRelatedPart(string relationshipType) {
            foreach (LegacyPptRelationshipView relationship in PackageRelationships) {
                if (relationship.IsType(relationshipType) && !relationship.TargetPart.IsNone) return relationship.TargetPart;
            }
            return default;
        }

        /// <summary>Every part reachable from the package relationships, each once, in relationship order.</summary>
        internal IReadOnlyList<LegacyPptPartView> ReachableParts() {
            var parts = new List<LegacyPptPartView>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<LegacyPptPartView>();
            for (int index = PackageRelationships.Count - 1; index >= 0; index--) {
                if (!PackageRelationships[index].TargetPart.IsNone) pending.Push(PackageRelationships[index].TargetPart);
            }
            while (pending.Count > 0) {
                LegacyPptPartView part = pending.Pop();
                if (!seen.Add(part.Name)) continue;
                parts.Add(part);
                IReadOnlyList<LegacyPptRelationshipView> relationships = part.Relationships;
                for (int index = relationships.Count - 1; index >= 0; index--) {
                    if (!relationships[index].TargetPart.IsNone) pending.Push(relationships[index].TargetPart);
                }
            }
            return parts;
        }

        /// <summary>Slide parts in the order of the presentation's slide list.</summary>
        internal IReadOnlyList<LegacyPptPartView> SlideParts => ListedParts("sldIdLst", "sldId");

        /// <summary>Slide master parts in the order of the presentation's master list.</summary>
        internal IReadOnlyList<LegacyPptPartView> SlideMasterParts => ListedParts("sldMasterIdLst", "sldMasterId");

        internal LegacyPptPartView NotesMaster => Presentation.RelatedPart(NotesMasterRelationshipType);

        internal LegacyPptPartView HandoutMaster => Presentation.RelatedPart(HandoutMasterRelationshipType);

        private IReadOnlyList<LegacyPptPartView> ListedParts(string listName, string entryName) {
            var parts = new List<LegacyPptPartView>();
            OfficeXmlNode list = Presentation.Root.Child(OfficeXmlNamespaces.PresentationML, listName);
            foreach (OfficeXmlNode entry in list.Children(OfficeXmlNamespaces.PresentationML, entryName)) {
                LegacyPptPartView part = Presentation.PartById(entry.Attribute(OfficeXmlNamespaces.Relationships, "id"));
                if (!part.IsNone) parts.Add(part);
            }
            return parts;
        }

        /// <summary>
        /// The element of this copy at the same place as <paramref name="element"/> in the document's part
        /// <paramref name="partName"/>, read from the document in either model; none when this copy has nothing there.
        /// </summary>
        internal OfficeXmlNode Resolve(string partName, OfficeXmlNode element) {
            if (element.IsNone) return default;
            OfficeXmlNode root = Part(partName).Root;
            if (root.IsNone) return default;
            var path = new Stack<int>();
            OfficeXmlNode current = element;
            while (true) {
                OfficeXmlNode parent = current.Parent;
                if (parent.IsNone) break;
                int index = 0;
                for (OfficeXmlNode sibling = current.PreviousSibling; !sibling.IsNone; sibling = sibling.PreviousSibling) index++;
                path.Push(index);
                current = parent;
            }
            if (!string.Equals(current.LocalName, root.LocalName, StringComparison.Ordinal)) return default;
            OfficeXmlNode resolved = root;
            while (path.Count > 0) {
                int index = path.Pop();
                OfficeXmlNode child = resolved.FirstChild;
                for (int position = 0; position < index && !child.IsNone; position++) child = child.NextSibling;
                if (child.IsNone) return default;
                resolved = child;
            }
            return string.Equals(resolved.LocalName, element.LocalName, StringComparison.Ordinal) ? resolved : default;
        }

        internal OfficeXmlNode RootOf(string partName) {
            if (!IsXml(partName)) return default;
            OfficeXmlTree tree = _package.GetXml(partName);
            return new OfficeXmlNode(tree, tree.Root);
        }

        internal string? ContentTypeOf(string partName) => _package.GetContentType(partName);

        internal byte[] BytesOf(string partName) => _package.ReadPartBytes(partName);

        internal bool TryGetStoredLength(string partName, out long length) => _package.TryGetStoredLength(partName, out length);

        private bool IsXml(string partName) {
            string? contentType = _package.GetContentType(partName);
            if (contentType == null) return partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
            return contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, "application/xml", StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, "text/xml", StringComparison.OrdinalIgnoreCase);
        }

        internal IReadOnlyList<LegacyPptRelationshipView> RelationshipsOf(string? sourcePartName) {
            string key = sourcePartName ?? string.Empty;
            if (_relationships.TryGetValue(key, out IReadOnlyList<LegacyPptRelationshipView>? cached)) return cached;
            var relationships = new List<LegacyPptRelationshipView>();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
            if (_package.PartExists(relationshipsPart)) {
                OfficeXmlTree tree = _package.GetXml(relationshipsPart);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                    node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    string id = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id") ?? string.Empty;
                    string type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty;
                    string target = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target") ?? string.Empty;
                    bool external = string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                    LegacyPptPartView targetPart = default;
                    if (!external) {
                        string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, target);
                        targetPart = Part(resolved);
                    }
                    relationships.Add(new LegacyPptRelationshipView(id, type, target, external, targetPart));
                }
            }
            _relationships[key] = relationships;
            return relationships;
        }

        private LegacyPptPartView FindPresentation() => PackageRelatedPart(OfficeDocumentRelationshipType);

        public void Dispose() => _package.Dispose();
    }

    /// <summary>A part of a <see cref="LegacyPptPackageView"/>.</summary>
    internal readonly struct LegacyPptPartView : IEquatable<LegacyPptPartView> {
        private const string SlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        private const string SlideLayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        private const string SlideMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        private const string NotesSlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        private const string NotesMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
        private const string HandoutMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
        private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
        private const string ThemeOverrideContentType = "application/vnd.openxmlformats-officedocument.themeOverride+xml";

        private readonly LegacyPptPackageView? _package;

        internal LegacyPptPartView(LegacyPptPackageView package, string name) {
            _package = package;
            Name = name;
        }

        internal bool IsNone => _package == null;

        /// <summary>The part name, such as <c>/ppt/slides/slide1.xml</c>; empty for none.</summary>
        internal string Name { get; } = string.Empty;

        internal LegacyPptPackageView? Package => _package;

        internal string? ContentType => _package?.ContentTypeOf(Name);

        /// <summary>The root element, or none for a part that is not XML.</summary>
        internal OfficeXmlNode Root => _package == null ? default : _package.RootOf(Name);

        /// <summary>The part's bytes as the package holds them.</summary>
        internal byte[] GetBytes() => _package == null ? Array.Empty<byte>() : _package.BytesOf(Name);

        internal IReadOnlyList<LegacyPptRelationshipView> Relationships =>
            _package == null ? Array.Empty<LegacyPptRelationshipView>() : _package.RelationshipsOf(Name);

        internal bool IsSlide => HasContentType(SlideContentType);
        internal bool IsSlideLayout => HasContentType(SlideLayoutContentType);
        internal bool IsSlideMaster => HasContentType(SlideMasterContentType);
        internal bool IsNotesSlide => HasContentType(NotesSlideContentType);
        internal bool IsNotesMaster => HasContentType(NotesMasterContentType);
        internal bool IsHandoutMaster => HasContentType(HandoutMasterContentType);
        internal bool IsTheme => HasContentType(ThemeContentType);
        internal bool IsThemeOverride => HasContentType(ThemeOverrideContentType);

        internal bool HasContentType(string contentType) => string.Equals(ContentType, contentType, StringComparison.OrdinalIgnoreCase);

        /// <summary>The relationship with this id, if the part has it.</summary>
        internal bool TryGetRelationship(string? id, out LegacyPptRelationshipView relationship) {
            if (id != null) {
                foreach (LegacyPptRelationshipView candidate in Relationships) {
                    if (!string.Equals(candidate.Id, id, StringComparison.Ordinal)) continue;
                    relationship = candidate;
                    return true;
                }
            }
            relationship = default;
            return false;
        }

        /// <summary>The internal part a relationship id targets, or none.</summary>
        internal LegacyPptPartView PartById(string? id) => TryGetRelationship(id, out LegacyPptRelationshipView relationship) ? relationship.TargetPart : default;

        /// <summary>The part the first internal relationship of this type targets, or none.</summary>
        internal LegacyPptPartView RelatedPart(string relationshipType) {
            foreach (LegacyPptRelationshipView relationship in Relationships) {
                if (relationship.IsType(relationshipType) && !relationship.TargetPart.IsNone) return relationship.TargetPart;
            }
            return default;
        }

        /// <summary>The parts the internal relationships of this type target, in relationship order.</summary>
        internal List<LegacyPptPartView> RelatedParts(string relationshipType) {
            var parts = new List<LegacyPptPartView>();
            foreach (LegacyPptRelationshipView relationship in Relationships) {
                if (relationship.IsType(relationshipType) && !relationship.TargetPart.IsNone) parts.Add(relationship.TargetPart);
            }
            return parts;
        }

        /// <summary>The relationship id under which this part relates to <paramref name="target"/>, or null.</summary>
        internal string? GetIdOfPart(LegacyPptPartView target) {
            foreach (LegacyPptRelationshipView relationship in Relationships) {
                if (relationship.TargetPart.Equals(target)) return relationship.Id;
            }
            return null;
        }

        /// <summary>For a slide master, its layouts in the order of its layout list.</summary>
        internal List<LegacyPptPartView> SlideLayoutParts() {
            var parts = new List<LegacyPptPartView>();
            OfficeXmlNode list = Root.Child(OfficeXmlNamespaces.PresentationML, "sldLayoutIdLst");
            foreach (OfficeXmlNode entry in list.Children(OfficeXmlNamespaces.PresentationML, "sldLayoutId")) {
                LegacyPptPartView part = PartById(entry.Attribute(OfficeXmlNamespaces.Relationships, "id"));
                if (!part.IsNone) parts.Add(part);
            }
            return parts;
        }

        internal LegacyPptPartView SlideLayout => RelatedPart(LegacyPptPackageView.SlideLayoutRelationshipType);
        internal LegacyPptPartView SlideMaster => RelatedPart(LegacyPptPackageView.SlideMasterRelationshipType);
        internal LegacyPptPartView NotesSlide => RelatedPart(LegacyPptPackageView.NotesSlideRelationshipType);
        internal LegacyPptPartView Theme => RelatedPart(LegacyPptPackageView.ThemeRelationshipType);
        internal LegacyPptPartView ThemeOverride => RelatedPart(LegacyPptPackageView.ThemeOverrideRelationshipType);

        public bool Equals(LegacyPptPartView other) =>
            ReferenceEquals(_package, other._package) && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);

        public override bool Equals(object? obj) => obj is LegacyPptPartView other && Equals(other);

        public override int GetHashCode() => StringComparer.OrdinalIgnoreCase.GetHashCode(Name);

        public override string ToString() => Name;
    }

    /// <summary>A relationship of a <see cref="LegacyPptPartView"/>.</summary>
    internal readonly struct LegacyPptRelationshipView {
        internal LegacyPptRelationshipView(string id, string type, string target, bool external, LegacyPptPartView targetPart) {
            Id = id;
            Type = type;
            Target = target;
            IsExternal = external;
            TargetPart = targetPart;
        }

        internal string Id { get; }

        /// <summary>The relationship type as written (a Strict type stays Strict; compare with <see cref="IsType"/>).</summary>
        internal string Type { get; }

        /// <summary>The target as written.</summary>
        internal string Target { get; }

        internal bool IsExternal { get; }

        /// <summary>The internal target part, or none for an external relationship or a missing part.</summary>
        internal LegacyPptPartView TargetPart { get; }

        /// <summary>Whether the relationship has this (transitional) type, written as Strict or transitional.</summary>
        internal bool IsType(string relationshipType) =>
            string.Equals(Type, relationshipType, StringComparison.Ordinal)
            || string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(Type), relationshipType, StringComparison.Ordinal);
    }
}
