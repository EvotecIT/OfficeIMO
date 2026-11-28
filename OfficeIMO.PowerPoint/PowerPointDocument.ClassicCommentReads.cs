using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int ClassicCommentNamespace = OfficeXmlNamespaces.PresentationML;
        private const string ClassicCommentsRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments";
        private const string ClassicAuthorsRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "commentAuthors";
        private (OfficeXmlTree Relationships, int Version, OfficeXmlNode Root)? _classicCommentAuthors;
        private (OfficeOpcEditablePackage Package, bool Readable)? _classicCommentReadPackage;

        private IReadOnlyList<PowerPointClassicComment>? TryReadClassicComments(PowerPointSlide slide) {
            if (ReadablePackage is not { } package || !CanReadClassicCommentsInPackage(package)) return null;
            if (slide.OwnerDocument != this || slide.RemovedFromPackage) {
                throw new ArgumentException("The slide does not belong to this presentation.", nameof(slide));
            }
            if (slide.PackagePartName is not { } slidePart || GetPackageSlideEntry(slidePart) == null) return null;
            if (!TryGetClassicCommentPart(package, slidePart, ClassicCommentsRelationship, "cmLst",
                out string? partName, out OfficeXmlNode root)) return null;
            if (partName == null) return Array.Empty<PowerPointClassicComment>();
            var comments = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode comment in root.Children(ClassicCommentNamespace, "cm")) {
                if (!HasTransitionalClassicCommentName(comment)) return null;
                foreach (OfficeXmlNode child in comment.Children()) {
                    if ((child.Is(ClassicCommentNamespace, "pos") || child.Is(ClassicCommentNamespace, "text")) &&
                        !HasTransitionalClassicCommentName(child)) return null;
                }
                comments.Add(comment);
            }
            if (comments.Count == 0) return Array.Empty<PowerPointClassicComment>();
            var binding = new PowerPointClassicCommentListBinding(this, slide, partName, comments);
            var result = new PowerPointClassicComment[comments.Count];
            for (int index = 0; index < result.Length; index++) {
                result[index] = new PowerPointClassicComment(this, slide, binding, index);
            }
            return result;
        }

        private bool CanReadClassicCommentsInPackage(OfficeOpcEditablePackage package) {
            if (_classicCommentReadPackage is { } checkedPackage && checkedPackage.Package == package) return checkedPackage.Readable;
            // The editable ZIP catalog survives commits that release the original read-only package.
            bool readable = _loadedPackage is { } loaded
                ? !loaded.HasDuplicateContentTypes && !loaded.HasInterleavedParts
                : !package.PartNames.Any(OfficeOpcPieceAssembler.HasPieceSuffix) && HasUniqueContentTypeEntries(package.ContentTypes);
            _classicCommentReadPackage = (package, readable);
            return readable;
        }

        internal bool TryReadClassicCommentAuthor(uint? authorId, out PowerPointCommentAuthor? author) {
            author = null;
            if (ReadablePackage is not { } package || !TryGetClassicAuthorRoot(package, out OfficeXmlNode authors)) return false;
            foreach (OfficeXmlNode entry in authors.Children(ClassicCommentNamespace, "cmAuthor")) {
                string? id = entry.Attribute("id");
                if (!HasTransitionalClassicCommentName(entry) || id == string.Empty) return false;
                if (id != null ? XmlConvert.ToUInt32(id) != authorId : authorId.HasValue) continue;
                author = PowerPointCommentAuthor.FromImportedClassic(entry.Attribute("name") ?? "Unknown", entry.Attribute("initials"));
                return true;
            }
            author = new PowerPointCommentAuthor("Unknown");
            return true;
        }

        private static bool HasTransitionalClassicCommentName(OfficeXmlNode element) =>
            element.Tree!.UriIdOf(element.Node) == OfficeXmlNamespaces.TransitionalUriId(ClassicCommentNamespace);

        private bool TryGetClassicAuthorRoot(OfficeOpcEditablePackage package, out OfficeXmlNode root) {
            root = default;
            if ((_packageSlideEntries?.Main ?? FirstRelated(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _)) is not { } main) return false;
            string partName = OfficeOpcPartNames.GetRelationshipsPartName(main);
            if (!package.PartExists(partName)) return true;
            if (GetPackageXml(partName) is not { } relationships) return false;
            if (_classicCommentAuthors is { } cached && cached.Relationships == relationships && cached.Version == relationships.Version) {
                root = cached.Root;
                return true;
            }

            // The parsed relationship list is shared by every author's lookup until this presentation's links change.
            if (!TryGetClassicCommentPart(package, main, ClassicAuthorsRelationship, "cmAuthorLst", out _, out root)) return false;
            _classicCommentAuthors = (relationships, relationships.Version, root);
            return true;
        }

        private bool TryGetClassicCommentPart(OfficeOpcEditablePackage package, string owner, string relationship, string rootName,
            out string? partName, out OfficeXmlNode root) {
            partName = null;
            root = default;
            try {
                IReadOnlyList<OfficeOpcRelationship>? edges = ReadEditablePackageRelationships(package, owner, strictXml: true, requireTransitionalTypes: true);
                if (edges == null) return false;
                foreach (OfficeOpcRelationship edge in edges) {
                    if (edge.Type != relationship || edge.IsExternal) continue;
                    if (partName != null || string.Equals(edge.Target, "NULL", StringComparison.OrdinalIgnoreCase) ||
                        edge.TargetPartName is not { } target || package.GetPartName(target) is not { } stored ||
                        stored.IndexOf('%') >= 0 || target.IndexOf('%') >= 0 ||
                        package.GetContentType(stored) is not { } contentType || !MatchesSubtypeSavePart(relationship, contentType) ||
                        GetPackageXml(stored) is not { } tree) return false;
                    root = new OfficeXmlNode(tree, tree.Root);
                    if (!root.Is(ClassicCommentNamespace, rootName) || !HasTransitionalClassicCommentName(root)) return false;
                    partName = stored;
                }
                return true;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException or FormatException) {
                return false;
            }
        }
    }
}
