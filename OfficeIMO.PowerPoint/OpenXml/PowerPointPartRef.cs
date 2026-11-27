using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXmlSdk;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.OpenXml {
    /// <summary>What a part is to the presentation, as far as content in it cares.</summary>
    internal enum PowerPointPartKind {
        Other,
        Slide,
        SlideLayout,
        SlideMaster,
        NotesSlide,
        NotesMaster
    }

    /// <summary>A relationship of a part: its id, type and target, and the part it targets when it is internal.</summary>
    internal readonly struct PowerPointRelationship {
        internal PowerPointRelationship(string id, string type, string target, bool external, PowerPointPartRef targetPart, Uri? uri = null) {
            Id = id;
            Type = type;
            Target = target;
            IsExternal = external;
            TargetPart = targetPart;
            Uri = uri;
        }

        /// <summary>
        /// The target as the SDK holds it, for an SDK relationship: a target the runtime can't parse is stored under a
        /// placeholder, and this URI still reads back as the original text.
        /// </summary>
        internal Uri? Uri { get; }

        internal string Id { get; }
        internal string Type { get; }

        /// <summary>The target as written: a URI for an external relationship, a reference relative to the source otherwise.</summary>
        internal string Target { get; }

        internal bool IsExternal { get; }

        /// <summary>The internal target part, or none for an external relationship or a target that is not a part.</summary>
        internal PowerPointPartRef TargetPart { get; }
    }

    /// <summary>
    /// A part of the presentation, wherever it is held: a part name in the document's package while the document works
    /// without the Open XML SDK, or the SDK part after the SDK opened. Content reads and edits its part's relationships
    /// through this type only; the SDK side lives in the bridge.
    /// </summary>
    internal readonly struct PowerPointPartRef : IEquatable<PowerPointPartRef> {
        internal const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        internal const string SlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        internal const string NotesSlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
        internal const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        internal const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        internal const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        internal const string ThemeOverrideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
        private const string SlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        private const string LayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        private const string MasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        private const string NotesContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        private const string NotesMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";

        private readonly PowerPointDocument? _document;
        private readonly string? _name;
        private readonly object? _sdk;

        private PowerPointPartRef(PowerPointDocument document, string name) {
            _document = document;
            _name = name;
            _sdk = null;
        }

        private PowerPointPartRef(object part) {
            _document = null;
            _name = null;
            _sdk = part;
        }

        /// <summary>A part of the document's package.</summary>
        internal static PowerPointPartRef Package(PowerPointDocument document, string partName) => new(document, partName);

        /// <summary>An SDK part, passed as <see cref="object"/> so callers name no SDK type.</summary>
        internal static PowerPointPartRef FromSdk(object? part) => part == null ? default : new PowerPointPartRef(part);

        internal static PowerPointPartRef None => default;

        internal bool IsNone => _name == null && _sdk == null;

        internal bool IsSdk => _sdk != null;

        /// <summary>The SDK part, or null for a part of the package.</summary>
        internal object? SdkPart => _sdk;

        /// <summary>The document, for a part of its package or an SDK part opened by it.</summary>
        internal PowerPointDocument? Document => _document ?? (_sdk != null ? PowerPointSdkParts.DocumentOf(_sdk) : null);

        /// <summary>
        /// The same part as it is held now: once the SDK opened the package, a part of the package is its SDK part.
        /// </summary>
        internal PowerPointPartRef Current =>
            _document is { IsSdkLoaded: true } document && _name != null ? FromSdk(document.FindSdkPart(_name)) : this;

        /// <summary>The part name, such as <c>/ppt/slides/slide1.xml</c>.</summary>
        internal string Name => _name ?? (_sdk != null ? PowerPointSdkParts.Name(_sdk) : string.Empty);

        internal PowerPointPartKind Kind {
            get {
                PowerPointPartRef current = Current;
                if (current._sdk != null) return PowerPointSdkParts.Kind(current._sdk);
                return current._document?.ReadablePackage?.GetContentType(current._name!) switch {
                    SlideContentType => PowerPointPartKind.Slide,
                    LayoutContentType => PowerPointPartKind.SlideLayout,
                    MasterContentType => PowerPointPartKind.SlideMaster,
                    NotesContentType => PowerPointPartKind.NotesSlide,
                    NotesMasterContentType => PowerPointPartKind.NotesMaster,
                    _ => PowerPointPartKind.Other
                };
            }
        }

        /// <summary>The part's root element, or none when the part has no XML.</summary>
        internal OfficeXmlNode Root {
            get {
                PowerPointPartRef current = Current;
                if (current._sdk != null) return OfficeXmlNode.FromSdk(PowerPointSdkParts.Root(current._sdk));
                OfficeXmlTree? tree = current._document?.GetPackageXml(current._name!);
                if (tree != null) return new OfficeXmlNode(tree, tree.Root);
                // A part over the size the package reads is read through the SDK, which applies the load's own limit.
                if (current._document is { } document && document.IsPackagePartOverLimit(current._name!)) {
                    document.OpenSdkForOversizedPart();
                    return Current.Root;
                }
                return default;
            }
        }

        /// <summary>Records an edit of the part's content, so the document writes it.</summary>
        internal void MarkEdited() {
            PowerPointPartRef current = Current;
            // An SDK part's edit is recorded too, so a read-only copy writes the parts that changed.
            if (current._sdk != null) current.Document?.MarkSdkPartEdited(current.Name);
            else current._document?.MarkPackageEdited();
        }

        // Records a relationship change, so cached edges between parts are read again.
        private void MarkRelationshipsEdited() {
            MarkEdited();
            Current._document?.InvalidatePackageRelationships();
        }

        // ---- Relationships ------------------------------------------------------------------------------------

        internal bool TryGetRelationship(string id, out PowerPointRelationship relationship) {
            relationship = default;
            if (string.IsNullOrEmpty(id)) return false;
            PowerPointPartRef current = Current;
            if (current._sdk != null) {
                if (!PowerPointSdkParts.TryGetRelationship(current._sdk, id, out string type, out string target, out bool external, out object? targetPart, out Uri? uri)) return false;
                relationship = new PowerPointRelationship(id, type, target, external, FromSdk(targetPart), uri);
                return true;
            }
            if (current._document?.ReadablePackage is not { } package) return false;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (!package.PartExists(relationshipsPart)) return false;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            int node = FindRelationship(tree, id);
            if (node == OfficeXmlTree.NoNode) return false;
            string relationshipTarget = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target") ?? string.Empty;
            bool isExternal = string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            string? partName = isExternal ? null : OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationshipTarget);
            relationship = new PowerPointRelationship(id, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty, relationshipTarget, isExternal,
                partName != null && package.PartExists(partName) ? Package(current._document, partName) : default);
            return true;
        }

        /// <summary>The part the first internal relationship of <paramref name="relationshipType"/> targets, or none.</summary>
        internal PowerPointPartRef RelatedPart(string relationshipType) {
            foreach (PowerPointPartRef part in RelatedParts(relationshipType)) return part;
            return default;
        }

        /// <summary>
        /// The parts the internal relationships of <paramref name="relationshipType"/> target, in relationship order.
        /// Strict relationship types match their transitional names.
        /// </summary>
        internal IEnumerable<PowerPointPartRef> RelatedParts(string relationshipType) {
            PowerPointPartRef current = Current;
            if (current._sdk != null) {
                foreach (object part in PowerPointSdkParts.RelatedParts(current._sdk, relationshipType)) yield return FromSdk(part);
                yield break;
            }
            if (current._document?.ReadablePackage is not { } package) yield break;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (!package.PartExists(relationshipsPart)) yield break;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                if (type == null || !string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(type), relationshipType, StringComparison.Ordinal)) continue;
                if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? partName = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target") ?? string.Empty);
                if (partName != null && package.PartExists(partName)) yield return Package(current._document, partName);
            }
        }

        /// <summary>
        /// Points the part's relationships of <paramref name="relationshipType"/> at <paramref name="target"/> instead: the old
        /// ones go and one new relationship, with the SDK's kind of id, is added.
        /// </summary>
        internal void ReplaceRelatedPart(string relationshipType, PowerPointPartRef target) {
            PowerPointPartRef current = Current, related = target.Current;
            if (current._sdk != null) {
                PowerPointSdkParts.ReplaceRelatedPart(current._sdk, relationshipType, related._sdk!);
                return;
            }
            OfficeOpcEditablePackage package = current.EditablePackage();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (package.PartExists(relationshipsPart)) {
                OfficeXmlTree tree = package.GetXml(relationshipsPart);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;) {
                    int next = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                    string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                    if (type != null && string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(type), relationshipType, StringComparison.Ordinal)) tree.Remove(node);
                    node = next;
                }
            }
            package.AddRelationship(current._name, relationshipType, OfficeOpcPartNames.GetRelativeTarget(current._name!, related._name!),
                id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));
            current.MarkRelationshipsEdited();
        }

        /// <summary>The id of the first internal relationship to <paramref name="target"/>, or null.</summary>
        internal string? RelationshipIdOf(PowerPointPartRef target) {
            PowerPointPartRef current = Current, wanted = target.Current;
            if (current._sdk != null) return wanted._sdk == null ? null : PowerPointSdkParts.IdOfPart(current._sdk, wanted._sdk);
            if (current._document?.ReadablePackage is not { } package) return null;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (!package.PartExists(relationshipsPart)) return null;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? partName = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target") ?? string.Empty);
                if (partName != null && Package(current._document, partName).Equals(wanted)) return tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
            }
            return null;
        }

        /// <summary>Adds an external hyperlink relationship to <paramref name="uri"/> and returns its id.</summary>
        internal string AddExternalHyperlink(Uri uri) {
            PowerPointPartRef current = Current;
            if (current._sdk != null) return PowerPointSdkParts.AddHyperlink(current._sdk, uri, external: true);
            string id = current.EditablePackage().AddRelationship(current._name, HyperlinkRelationshipType, PowerPointHyperlinkResolver.TargetText(uri), external: true);
            current.MarkRelationshipsEdited();
            return id;
        }

        /// <summary>
        /// The id of an internal hyperlink relationship to <paramref name="target"/>, written as a relative reference,
        /// reusing one that already points there.
        /// </summary>
        internal string GetOrAddRelativeHyperlink(PowerPointPartRef target) {
            PowerPointPartRef current = Current;
            target = target.Current;
            if (current._sdk != null) return PowerPointSdkParts.GetOrAddRelativeHyperlink(current._sdk, target._sdk ?? throw new InvalidOperationException("The target part is not open in the SDK."));
            OfficeOpcEditablePackage package = current.EditablePackage();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (package.PartExists(relationshipsPart)) {
                OfficeXmlTree tree = package.GetXml(relationshipsPart);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                    node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (!string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type"), HyperlinkRelationshipType, StringComparison.Ordinal)
                        || string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target")), target.Name, StringComparison.Ordinal)
                        && tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } existing) {
                        return existing;
                    }
                }
            }
            string id = package.AddRelationship(current._name, HyperlinkRelationshipType, OfficeOpcPartNames.GetRelativeTarget(current._name!, target.Name));
            current.MarkRelationshipsEdited();
            return id;
        }

        /// <summary>The id of a part relationship to <paramref name="target"/>, adding a slide relationship when there is none.</summary>
        internal string GetOrAddSlideRelationship(PowerPointPartRef target) {
            PowerPointPartRef current = Current;
            target = target.Current;
            if (current._sdk != null) return PowerPointSdkParts.GetOrAddPartRelationship(current._sdk, target._sdk ?? throw new InvalidOperationException("The target part is not open in the SDK."));
            OfficeOpcEditablePackage package = current.EditablePackage();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(current._name);
            if (package.PartExists(relationshipsPart)) {
                OfficeXmlTree tree = package.GetXml(relationshipsPart);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                    node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    // A hyperlink relationship to the slide is not the part relationship the SDK would find.
                    if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type"), HyperlinkRelationshipType, StringComparison.Ordinal)
                        || string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target")), target.Name, StringComparison.Ordinal)
                        && tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } existing) {
                        return existing;
                    }
                }
            }
            string id = package.AddRelationship(current._name, SlideRelationshipType, OfficeOpcPartNames.GetRelativeTarget(current._name!, target.Name));
            current.MarkRelationshipsEdited();
            return id;
        }

        /// <summary>
        /// Removes a relationship nothing in the part refers to any more: an external one, or an internal one together
        /// with the media part only it used. A notes slide keeps its relationship to <paramref name="owningSlide"/>.
        /// </summary>
        internal void RemoveRelationshipIfUnused(string id, PowerPointPartRef owningSlide) {
            if (string.IsNullOrEmpty(id)) return;
            PowerPointPartRef current = Current;
            if (References(current.Root, id)) return;
            if (current._sdk != null) {
                PowerPointSdkParts.RemoveRelationshipIfUnused(current._sdk, id, owningSlide.Current._sdk);
                return;
            }
            if (!current.TryGetRelationship(id, out PowerPointRelationship relationship)) return;
            if (current.Kind == PowerPointPartKind.NotesSlide && !owningSlide.IsNone && relationship.TargetPart.Equals(owningSlide.Current)
                && !string.Equals(relationship.Type, HyperlinkRelationshipType, StringComparison.Ordinal)) return;
            // A read-only presentation can't be saved, so its relationships are left as they are.
            if (current._document?.EditablePackage is not { } package) return;
            OfficeXmlTree tree = package.GetXml(OfficeOpcPartNames.GetRelationshipsPartName(current._name));
            int node = FindRelationship(tree, id);
            if (node != OfficeXmlTree.NoNode) tree.Remove(node);
            current.MarkRelationshipsEdited();
            // Media used only through this relationship leaves with it; slides and other parts stay.
            if (relationship.TargetPart._name is { } target && IsMedia(package, target) && !IsTargeted(package, target)) package.DeletePart(target);
        }

        // Whether an r: attribute on the element or below it carries the relationship id.
        private static bool References(OfficeXmlNode element, string id) {
            if (element.IsNone) return false;
            if (element.IsSdk) return PowerPointSdkParts.ReferencesRelationship(element.SdkElement!, id);
            OfficeXmlTree tree = element.Tree!;
            var pending = new Stack<int>();
            pending.Push(element.Node);
            while (pending.Count > 0) {
                int node = pending.Pop();
                if (tree.HasRelationshipAttribute(node, id)) return true;
                for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                    if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) pending.Push(child);
                }
            }
            return false;
        }

        private static bool IsMedia(OfficeOpcEditablePackage package, string partName) {
            string? contentType = package.GetContentType(partName);
            return contentType != null && (contentType.StartsWith("audio/", StringComparison.OrdinalIgnoreCase) || contentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase));
        }

        // Whether any relationship in the package still targets the part.
        private static bool IsTargeted(OfficeOpcEditablePackage package, string partName) {
            foreach (string name in package.PartNames.ToList()) {
                if (!name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                OfficeXmlTree tree = package.GetXml(name);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                    node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(OfficeOpcPartNames.ResolveRelationshipTarget(name, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target")), partName, StringComparison.Ordinal)) return true;
                }
            }
            return false;
        }

        private static int FindRelationship(OfficeXmlTree tree, string id) {
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id"), id, StringComparison.Ordinal)) return node;
            }
            return OfficeXmlTree.NoNode;
        }

        private OfficeOpcEditablePackage EditablePackage() =>
            _document?.EditablePackage ?? throw new InvalidOperationException("The presentation can't be edited.");

        /// <summary>The part's path relative to <paramref name="source"/>, as a relationship target from it is written.</summary>
        internal string RelativeTo(PowerPointPartRef source) => OfficeOpcPartNames.GetRelativeTarget(source.Name, Name);

        public bool Equals(PowerPointPartRef other) {
            PowerPointPartRef left = Current, right = other.Current;
            if (left._sdk != null || right._sdk != null) return ReferenceEquals(left._sdk, right._sdk);
            return ReferenceEquals(left._document, right._document) && string.Equals(left._name, right._name, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj) => obj is PowerPointPartRef other && Equals(other);

        public override int GetHashCode() => StringComparer.OrdinalIgnoreCase.GetHashCode(Name);

        public override string ToString() => Name;
    }
}
