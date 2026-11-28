using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string ClassicCommentContentType = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
        private const string ClassicAuthorContentType = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
        private HashSet<OfficeXmlTree>? _validatedClassicCommentParts;
        private Dictionary<OfficeXmlTree, (int Version, IReadOnlyList<OfficeOpcRelationship> Links, bool XmlValidated)>? _classicCommentMutationLinks;

        private void ResetClassicCommentAddValidation() {
            _validatedClassicCommentParts = null;
            _classicCommentMutationLinks = null;
            _classicCommentGraph = null;
        }

        private OfficeOpcEditablePackage? GetEditableClassicCommentPackage(PowerPointSlide slide) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(slide);
            if (EditablePackage is not { } package || !CanReadClassicCommentsInPackage(package)) return null;
            if (slide.OwnerDocument != this || slide.RemovedFromPackage) {
                throw new ArgumentException("The slide does not belong to this presentation.", nameof(slide));
            }
            return slide.PackagePartName is { } part && GetPackageSlideEntry(part) != null ? package : null;
        }

        private PowerPointClassicComment? TryAddPackageClassicComment(OfficeOpcEditablePackage package, PowerPointSlide slide,
            PowerPointCommentAuthor author, string text, long x, long y, DateTime? created) {
            if (created is { Kind: not DateTimeKind.Utc } || text.IndexOf('\r') >= 0) return null;
            string main;
            string slideName = slide.PackagePartName!;
            string? authorsPart;
            string? commentPart = null;
            string? newCommentsRelId = null;
            string? newAuthorsRelId = null;
            string? newCommentTarget = null;
            bool foundSlide = false;
            OfficeXmlNode authors, comments = default, matching = default;
            uint authorId;
            (uint Next, uint Last) indexes;
            try {
                if (_packageSlideEntries is not { } slides ||
                    !TryReadClassicCommentMutationPart(package, slides.Main, ClassicAuthorsRelationship, "cmAuthorLst",
                        out var mainLinks, out authorsPart, out authors)) return null;
                main = slides.Main;
                var usedAuthorIds = new HashSet<uint>();
                foreach (OfficeXmlNode entry in authors.Children(ClassicCommentNamespace, "cmAuthor")) {
                    if (!HasTransitionalClassicCommentName(entry)) return null;
                    uint? id = ReadClassicUnsigned(entry, "id");
                    if (id.HasValue) usedAuthorIds.Add(id.Value);
                    if (matching.IsNone && entry.Attribute("name") == author.Name && entry.Attribute("initials") == author.ClassicInitials) {
                        matching = entry;
                    }
                }
                if (matching.IsNone) {
                    authorId = FindAvailableUInt32Id(usedAuthorIds, 0U, "classic-comment-author");
                } else {
                    uint? found = ReadClassicUnsigned(matching, "id");
                    if (!found.HasValue) return null;
                    authorId = found.Value;
                }

                var usedIndexes = new HashSet<uint>();
                var linkedParts = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
                IReadOnlyList<OfficeOpcRelationship>? targetLinks = null;
                foreach ((_, _, string part) in slides.Entries) {
                    if (!TryReadClassicCommentMutationPart(package, part, ClassicCommentsRelationship, "cmLst",
                        out var links, out string? entryPart, out OfficeXmlNode entries)) return null;
                    if (entryPart != null) linkedParts.Add(entryPart);
                    if (PartIdentity(part).Equals(PartIdentity(slideName), StringComparison.OrdinalIgnoreCase)) {
                        foundSlide = true;
                        targetLinks = links;
                        commentPart = entryPart;
                        comments = entries;
                    }
                    foreach (OfficeXmlNode entry in entries.Children(ClassicCommentNamespace, "cm")) {
                        if (!HasTransitionalClassicCommentName(entry)) return null;
                        if (ReadClassicUnsigned(entry, "authorId") == authorId && ReadClassicUnsigned(entry, "idx") is { } index) {
                            usedIndexes.Add(index);
                        }
                    }
                }
                if (!foundSlide || targetLinks == null ||
                    package.CountParts(ClassicAuthorContentType) != (authorsPart == null ? 0 : 1) ||
                    package.CountParts(ClassicCommentContentType) != linkedParts.Count) return null;
                if (authorsPart == null) newAuthorsRelId = ReservePackageMutationRelationshipId(mainLinks!);
                if (commentPart == null) {
                    newCommentsRelId = ReservePackageMutationRelationshipId(targetLinks);
                    newCommentTarget = OfficeOpcPartNames.ResolveRelationshipTarget(OfficeOpcPartNames.GetRelationshipsPartName(slideName),
                        "../comments/comment.xml");
                    if (newCommentTarget == null) return null;
                }
                indexes = FindClassicCommentIndexes(usedIndexes, matching.IsNone ? null : ReadClassicUnsigned(matching, "lastIdx"));
            } catch (Exception exception) when (exception is XmlException or InvalidDataException or FormatException or OverflowException) {
                return null;
            }

            ValidateClassicCommentAuthor(author);
            if (newAuthorsRelId != null) {
                string directory = main.Substring(0, main.LastIndexOf('/') + 1);
                authorsPart = UniquePartName(package, ClassicAuthorContentType, directory + "commentAuthors.xml");
            }
            if (newCommentsRelId != null) {
                string target = newCommentTarget!;
                commentPart = NextPartName(package, ClassicCommentContentType, target.Substring(0, target.Length - 4), ".xml");
            }

            if (newAuthorsRelId != null) {
                authors = CreateClassicCommentPart(package, main, authorsPart!, ClassicAuthorContentType,
                    ClassicAuthorsRelationship, "cmAuthorLst", newAuthorsRelId);
            }
            if (matching.IsNone) {
                matching = authors.CreateElement(ClassicCommentNamespace, "cmAuthor");
                matching.SetUInt32Attribute("id", authorId);
                matching.SetAttribute("name", author.Name);
                matching.SetAttribute("initials", author.ClassicInitials);
                matching.SetUInt32Attribute("lastIdx", indexes.Last);
                matching.SetUInt32Attribute("clrIdx", authorId);
                authors.Append(matching);
            } else {
                matching.SetUInt32Attribute("lastIdx", indexes.Last);
            }
            if (newCommentsRelId != null) {
                comments = CreateClassicCommentPart(package, slideName, commentPart!, ClassicCommentContentType,
                    ClassicCommentsRelationship, "cmLst", newCommentsRelId);
            }

            OfficeXmlNode comment = comments.CreateElement(ClassicCommentNamespace, "cm");
            comment.SetUInt32Attribute("authorId", authorId);
            comment.SetAttribute("dt", XmlConvert.ToString(created ?? DateTime.UtcNow, XmlDateTimeSerializationMode.Utc));
            comment.SetUInt32Attribute("idx", indexes.Next);
            OfficeXmlNode position = comment.CreateElement(ClassicCommentNamespace, "pos");
            position.SetInt64Attribute("x", x);
            position.SetInt64Attribute("y", y);
            comment.Append(position);
            OfficeXmlNode content = comment.CreateElement(ClassicCommentNamespace, "text");
            content.SetInnerText(text);
            comment.Append(content);
            comments.Append(comment);

            if (newAuthorsRelId != null || newCommentsRelId != null) InvalidatePackageRelationships();
            MarkPackageEdited();
            var binding = new PowerPointClassicCommentListBinding(this, slide, commentPart!, new[] { comment });
            return new PowerPointClassicComment(this, slide, binding, 0);
        }

        private bool TryReadClassicCommentMutationPart(OfficeOpcEditablePackage package, string owner, string relationship,
            string rootName, out IReadOnlyList<OfficeOpcRelationship>? links, out string? partName, out OfficeXmlNode root) {
            partName = null;
            root = default;
            links = ReadClassicCommentMutationLinks(package, owner);
            return links != null && TryFindClassicCommentPart(package, links, relationship, rootName, out partName, out root) &&
                IsValidClassicCommentMutationPart(package, partName, root);
        }

        private static uint? ReadClassicUnsigned(OfficeXmlNode node, string attribute) =>
            node.Attribute(attribute) is { } value ? XmlConvert.ToUInt32(value) : null;

        private IReadOnlyList<OfficeOpcRelationship>? ReadClassicCommentMutationLinks(OfficeOpcEditablePackage package, string? owner,
            bool validateXml = false) {
            string part = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (!package.PartExists(part)) return Array.Empty<OfficeOpcRelationship>();
            if (package.GetContentType(part) != OfficeOpcEditablePackage.RelationshipContentType || GetPackageXml(part) is not { } tree) return null;
            var cached = default((int Version, IReadOnlyList<OfficeOpcRelationship> Links, bool XmlValidated));
            bool validCache = _classicCommentMutationLinks?.TryGetValue(tree, out cached) == true && cached.Version == tree.Version;
            if (validCache && (!validateXml || cached.XmlValidated)) return cached.Links;
            IReadOnlyList<OfficeOpcRelationship>? links = validCache ? cached.Links :
                ReadPackageMutationRelationships(package, owner, requireTransitionalTypes: true);
            if (links == null) return null;
            if (validateXml) {
                using Stream stream = package.OpenPartContentStream(part, default);
                PowerPointXmlReader.ValidatePackagePartXml(stream);
            }
            _classicCommentMutationLinks ??= new Dictionary<OfficeXmlTree, (int, IReadOnlyList<OfficeOpcRelationship>, bool)>();
            _classicCommentMutationLinks[tree] = (tree.Version, links, validateXml);
            return links;
        }

        private bool IsValidClassicCommentMutationPart(OfficeOpcEditablePackage package, string? partName, OfficeXmlNode root) {
            if (partName == null) return true;
            if (root.Tree is not { } tree) return false;
            if (_validatedClassicCommentParts?.Contains(tree) == true) return true;
            if (tree.HasCommentsOrInstructions()) return false;
            using Stream part = package.OpenPartContentStream(partName, default);
            PowerPointXmlReader.ValidatePackagePartXml(part);
            _validatedClassicCommentParts ??= new HashSet<OfficeXmlTree>();
            _validatedClassicCommentParts.Add(tree);
            return true;
        }

        private OfficeXmlNode CreateClassicCommentPart(OfficeOpcEditablePackage package, string owner, string part,
            string contentType, string relationship, string root, string id) {
            OfficeXmlTree tree = package.CreateXmlPart(part, contentType, ClassicCommentNamespace, root);
            package.AddRelationship(owner, relationship, part, id: id);
            EnsurePackageRelationshipContentType(package, owner);
            _validatedClassicCommentParts ??= new HashSet<OfficeXmlTree>();
            _validatedClassicCommentParts.Add(tree);
            return new OfficeXmlNode(tree, tree.Root);
        }
    }
}
