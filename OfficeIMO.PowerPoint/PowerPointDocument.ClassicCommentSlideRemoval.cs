using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private HashSet<string>? _validatedClassicCommentSlideXml;
        private Dictionary<OfficeXmlTree, int>? _validatedClassicCommentSlideTrees;

        private sealed record ClassicSlideCommentRemoval(List<(OfficeXmlNode Author, uint? LastIndex)> Authors,
            ClassicCommentPartDeletion? AuthorPart);

        private bool TryPlanPackageSlideClassicComments(OfficeOpcEditablePackage package, PackageSlideList slides, string slide,
            List<string> removed, out ClassicSlideCommentRemoval? plan) {
            plan = null;
            try {
                if (!CanReadClassicCommentsInPackage(package) || !ValidateSurvivingSlideXml(package, slides, slide) ||
                    !TryReadClassicCommentMutationPart(package, slide, ClassicCommentsRelationship, "cmLst",
                        out var slideLinks, out string? commentPart, out OfficeXmlNode comments) ||
                    commentPart == null || !removed.Contains(PartIdentity(commentPart), OfficeAsciiIgnoreCaseComparer.Instance)) return false;

                ClassicCommentPartDeletion? commentDeletion = PlanClassicCommentPartDeletion(
                    package, slide, commentPart, slideLinks!, ClassicCommentsRelationship);
                if (commentDeletion == null) return false;
                var affected = new HashSet<uint>();
                foreach (OfficeXmlNode comment in comments.Children(ClassicCommentNamespace, "cm")) {
                    if (!HasTransitionalClassicCommentName(comment)) return false;
                    if (ReadClassicUnsigned(comment, "authorId") is { } id) affected.Add(id);
                }

                if (affected.Count != 0 && !TryPlanSlideClassicCommentAuthors(package, slides, slide, affected, out plan)) return false;
                if (!CanMutateClassicCommentRelationships(package, commentDeletion, plan?.AuthorPart)) return false;
                // A slide's relationship part leaves even when it has links besides its comments part.
                return !removed.Any(part => _classicCommentGraph!.HasReference(OfficeOpcPartNames.GetRelationshipsPartName(part)));
            } catch (Exception exception) when (exception is XmlException or InvalidDataException or IOException or FormatException or
                ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }
        }

        private bool ValidateSurvivingSlideXml(OfficeOpcEditablePackage package, PackageSlideList slides, string removedSlide) {
            foreach ((_, _, string part) in slides.Entries) {
                if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(removedSlide, part)) continue;
                if (!ValidateSurvivingSlideXml(package, part)) return false;
            }
            return true;
        }

        private bool ValidateSurvivingSlideXml(OfficeOpcEditablePackage package, string part) {
            OfficeXmlTree? tree = package.GetParsedXml(part);
            if (tree != null && _validatedClassicCommentSlideTrees?.TryGetValue(tree, out int version) == true &&
                version == tree.Version) return true;

            // A replaced or newly created part is no longer a stored source member.
            bool stored = tree == null && package.TryGetStoredLength(part, out _);
            if (stored && _validatedClassicCommentSlideXml?.Contains(part) == true) return true;
            using Stream input = package.OpenPartContentStream(part, default);
            PowerPointXmlReader.ValidatePackagePartXml(input, "sld", OfficeXmlNamespaces.KnownUri(ClassicCommentNamespace));
            if (stored) {
                (_validatedClassicCommentSlideXml ??= new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance)).Add(part);
            } else if (tree != null) {
                (_validatedClassicCommentSlideTrees ??= new Dictionary<OfficeXmlTree, int>())[tree] = tree.Version;
            }
            return true;
        }

        private bool TryPlanSlideClassicCommentAuthors(OfficeOpcEditablePackage package, PackageSlideList slides, string slide,
            HashSet<uint> affected, out ClassicSlideCommentRemoval? plan) {
            plan = null;
            if (!TryReadClassicCommentMutationPart(package, slides.Main, ClassicAuthorsRelationship, "cmAuthorLst",
                    out var mainLinks, out string? authorPart, out OfficeXmlNode authors)) return false;

            var surviving = new Dictionary<uint, uint>();
            foreach ((_, _, string part) in slides.Entries) {
                if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(part), PartIdentity(slide))) continue;
                if (!TryReadClassicCommentMutationPart(package, part, ClassicCommentsRelationship, "cmLst",
                    out _, out _, out OfficeXmlNode comments)) return false;
                foreach (OfficeXmlNode comment in comments.Children(ClassicCommentNamespace, "cm")) {
                    if (!HasTransitionalClassicCommentName(comment)) return false;
                    if (ReadClassicUnsigned(comment, "authorId") is not { } id || !affected.Contains(id)) continue;
                    surviving.TryGetValue(id, out uint maximum);
                    surviving[id] = Math.Max(maximum, ReadClassicUnsigned(comment, "idx") ?? 0U);
                }
            }

            var edits = new List<(OfficeXmlNode, uint?)>();
            var matched = new HashSet<uint>();
            int remainingAuthors = 0;
            foreach (OfficeXmlNode author in authors.Children(ClassicCommentNamespace, "cmAuthor")) {
                if (!HasTransitionalClassicCommentName(author)) return false;
                remainingAuthors++;
                uint? id = ReadClassicUnsigned(author, "id");
                if (!id.HasValue || !affected.Contains(id.Value) || !matched.Add(id.Value)) continue;
                if (surviving.TryGetValue(id.Value, out uint lastIndex)) {
                    edits.Add((author, lastIndex));
                } else {
                    edits.Add((author, null));
                    remainingAuthors--;
                }
            }

            ClassicCommentPartDeletion? authorDeletion = null;
            if (authorPart != null && remainingAuthors == 0 && surviving.Count < affected.Count) {
                authorDeletion = PlanClassicCommentPartDeletion(package, slides.Main, authorPart, mainLinks!, ClassicAuthorsRelationship);
                if (authorDeletion == null) return false;
            }
            plan = new ClassicSlideCommentRemoval(edits, authorDeletion);
            return true;
        }

        private void ApplyPackageSlideClassicCommentRemoval(OfficeOpcEditablePackage package, ClassicSlideCommentRemoval plan) {
            foreach ((OfficeXmlNode author, uint? lastIndex) in plan.Authors) {
                if (lastIndex.HasValue) author.SetUInt32Attribute("lastIdx", lastIndex.Value);
                else author.Remove();
            }
            if (plan.AuthorPart != null) {
                DeleteClassicCommentPart(package, plan.AuthorPart);
                InvalidatePackageRelationships();
            }
        }
    }
}
