using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record ClassicCommentPartDeletion(string Owner, string Part, string Id,
            OfficeXmlTree Relationships, int Node, bool DeleteRelationships);

        internal bool TryRemovePackageClassicComment(PowerPointSlide slide, string partName, OfficeXmlNode comment) {
            if (GetEditableClassicCommentPackage(slide) is not { } package || _packageSlideEntries is not { } slides) return false;

            ClassicCommentPartDeletion? commentsDeletion = null, authorsDeletion = null;
            OfficeXmlNode author = default;
            uint? lastIndex = null;
            try {
                string slidePart = slide.PackagePartName!;
                OfficeXmlNode root = comment.Parent;
                if (!TryReadClassicCommentMutationPart(package, slidePart, ClassicCommentsRelationship, "cmLst",
                        out var slideLinks, out string? slideCommentsPart, out OfficeXmlNode slideComments) ||
                    slideCommentsPart != partName || !slideComments.Equals(root)) return false;
                if (root.Child(ClassicCommentNamespace, "cm").Equals(comment) &&
                    comment.NextSiblingNamed(ClassicCommentNamespace, "cm").IsNone) {
                    commentsDeletion = PlanClassicCommentPartDeletion(package, slidePart, partName, slideLinks!, ClassicCommentsRelationship);
                    if (commentsDeletion == null) return false;
                }

                if (ReadClassicUnsigned(comment, "authorId") is { } authorId) {
                    if (!TryReadClassicCommentMutationPart(package, slides.Main, ClassicAuthorsRelationship, "cmAuthorLst",
                        out var mainLinks, out string? authorsPart, out OfficeXmlNode authors)) return false;
                    int authorCount = 0;
                    foreach (OfficeXmlNode entry in authors.Children(ClassicCommentNamespace, "cmAuthor")) {
                        if (!HasTransitionalClassicCommentName(entry)) return false;
                        authorCount++;
                        if (author.IsNone && ReadClassicUnsigned(entry, "id") == authorId) author = entry;
                    }

                    foreach ((_, _, string part) in slides.Entries) {
                        if (!TryReadClassicCommentMutationPart(package, part, ClassicCommentsRelationship, "cmLst",
                            out _, out _, out OfficeXmlNode entries)) return false;
                        foreach (OfficeXmlNode entry in entries.Children(ClassicCommentNamespace, "cm")) {
                            if (!HasTransitionalClassicCommentName(entry)) return false;
                            if (entry.Equals(comment) || ReadClassicUnsigned(entry, "authorId") != authorId) continue;
                            uint index = ReadClassicUnsigned(entry, "idx") ?? 0U;
                            lastIndex = Math.Max(lastIndex.GetValueOrDefault(), index);
                        }
                    }

                    if (!lastIndex.HasValue && authorsPart != null && authorCount == (author.IsNone ? 0 : 1)) {
                        authorsDeletion = PlanClassicCommentPartDeletion(package, slides.Main, authorsPart, mainLinks!, ClassicAuthorsRelationship);
                        if (authorsDeletion == null) return false;
                    }
                }

                if (!CanMutateClassicCommentRelationships(package, commentsDeletion, authorsDeletion)) return false;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException or IOException or FormatException or
                ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }

            comment.Remove();
            if (!author.IsNone) {
                if (lastIndex.HasValue) author.SetUInt32Attribute("lastIdx", lastIndex.Value);
                else author.Remove();
            }
            if (commentsDeletion != null) DeleteClassicCommentPart(package, commentsDeletion);
            if (authorsDeletion != null) DeleteClassicCommentPart(package, authorsDeletion);
            if (commentsDeletion != null || authorsDeletion != null) InvalidatePackageRelationships();
            MarkPackageEdited();
            return true;
        }

        private static ClassicCommentPartDeletion? PlanClassicCommentPartDeletion(OfficeOpcEditablePackage package, string owner,
            string part, IReadOnlyList<OfficeOpcRelationship> links, string relationship) {
            // Deleting a part with relationships can also change reachability for its children.
            if (package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(part))) return null;
            string id = links.First(edge => !edge.IsExternal && edge.Type == relationship).Id;
            OfficeXmlTree relationships = package.GetRelationships(owner);
            int node = FindRelationship(relationships, id);
            return node == OfficeXmlTree.NoNode ? null : new ClassicCommentPartDeletion(owner, part, id, relationships, node, links.Count == 1);
        }

        private void DeleteClassicCommentPart(OfficeOpcEditablePackage package, ClassicCommentPartDeletion deletion) {
            string ownerRelationships = OfficeOpcPartNames.GetRelationshipsPartName(deletion.Owner);
            deletion.Relationships.Remove(deletion.Node);
            if (deletion.DeleteRelationships) package.DeletePart(ownerRelationships);
            package.DeletePart(deletion.Part);
            _classicCommentGraph!.RecordDeletion(deletion, ownerRelationships);
        }
    }
}
