using System.Xml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        internal bool TryReassignPackageClassicComment(PowerPointSlide slide, string partName, OfficeXmlNode comment,
            PowerPointCommentAuthor requestedAuthor) {
            if (GetEditableClassicCommentPackage(slide) is not { } package || _packageSlideEntries is not { } slides) return false;

            string? authorPart;
            string? authorRelationshipId = null;
            OfficeXmlNode authors, target, previous = default;
            uint targetId;
            (uint Next, uint Last) indexes;
            uint? survivingIndex = null;
            try {
                if (!TryReadClassicCommentMutationPart(package, slide.PackagePartName!, ClassicCommentsRelationship, "cmLst",
                        out _, out string? currentPart, out OfficeXmlNode commentList) ||
                    currentPart != partName || !commentList.Equals(comment.Parent) ||
                    !TryReadClassicCommentMutationPart(package, slides.Main, ClassicAuthorsRelationship, "cmAuthorLst",
                        out var mainLinks, out authorPart, out authors) ||
                    !TryFindClassicCommentAuthor(authors, requestedAuthor, out target, out targetId) ||
                    !CanMutateClassicCommentRelationships(package, comments: null, authors: null)) return false;

                uint? previousId = ReadClassicUnsigned(comment, "authorId");
                if (previousId == targetId) {
                    if (target.IsNone) return false;
                    ValidateClassicCommentAuthor(requestedAuthor);
                    return true;
                }

                var usedIndexes = new HashSet<uint>();
                foreach ((_, _, string part) in slides.Entries) {
                    if (!TryReadClassicCommentMutationPart(package, part, ClassicCommentsRelationship, "cmLst",
                        out _, out _, out OfficeXmlNode comments)) return false;
                    foreach (OfficeXmlNode entry in comments.Children(ClassicCommentNamespace, "cm")) {
                        if (!HasTransitionalClassicCommentName(entry)) return false;
                        if (entry.Equals(comment)) continue;
                        uint? id = ReadClassicUnsigned(entry, "authorId");
                        if (id == targetId) {
                            if (ReadClassicUnsigned(entry, "idx") is { } index) usedIndexes.Add(index);
                        } else if (previousId.HasValue && id == previousId) {
                            uint index = ReadClassicUnsigned(entry, "idx") ?? 0U;
                            survivingIndex = Math.Max(survivingIndex.GetValueOrDefault(), index);
                        }
                    }
                }

                if (previousId.HasValue) {
                    foreach (OfficeXmlNode entry in authors.Children(ClassicCommentNamespace, "cmAuthor")) {
                        if (ReadClassicUnsigned(entry, "id") != previousId) continue;
                        previous = entry;
                        break;
                    }
                }

                if (authorPart == null) {
                    if (package.CountParts(ClassicAuthorContentType) != 0) return false;
                    authorRelationshipId = ReservePackageMutationRelationshipId(mainLinks!);
                }
                indexes = FindClassicCommentIndexes(usedIndexes, target.IsNone ? null : ReadClassicUnsigned(target, "lastIdx"));
            } catch (Exception exception) when (exception is XmlException or InvalidDataException or IOException or FormatException or
                ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }

            ValidateClassicCommentAuthor(requestedAuthor);
            if (authorRelationshipId != null) {
                string directory = slides.Main.Substring(0, slides.Main.LastIndexOf('/') + 1);
                authorPart = UniquePartName(package, ClassicAuthorContentType, directory + "commentAuthors.xml");
                authors = CreateClassicCommentPart(package, slides.Main, authorPart, ClassicAuthorContentType,
                    ClassicAuthorsRelationship, "cmAuthorLst", authorRelationshipId);
                InvalidatePackageRelationships();
            }
            if (target.IsNone) AppendClassicCommentAuthor(authors, requestedAuthor, targetId, indexes.Last);
            else target.SetUInt32Attribute("lastIdx", indexes.Last);

            comment.SetUInt32Attribute("authorId", targetId);
            comment.SetUInt32Attribute("idx", indexes.Next);
            if (!previous.IsNone) {
                if (survivingIndex.HasValue) previous.SetUInt32Attribute("lastIdx", survivingIndex.Value);
                else previous.Remove();
            }
            MarkPackageEdited();
            return true;
        }
    }
}
