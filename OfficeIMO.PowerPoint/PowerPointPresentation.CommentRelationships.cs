using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private const string ModernCommentExtensionUri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";

        // A package relationship alone does not provide the explicit slide reference
        // defined by MS-PPTX 2.2.10. Keep that reference alongside the comment part.
        private static void EnsureModernCommentRelationship(SlidePart slide, PowerPointCommentPart part) {
            string relationshipId = slide.GetIdOfPart(part);
            P.Slide root = slide.Slide ?? throw new InvalidOperationException("The comment slide has no root element.");
            P.SlideExtensionList? extensions = root.SlideExtensionList;
            if (extensions == null) {
                extensions = new P.SlideExtensionList();
                root.AddChild(extensions, true);
            }
            P.SlideExtension[] commentExtensions = extensions.Elements<P.SlideExtension>()
                .Where(IsModernCommentExtension).ToArray();
            if (commentExtensions.SelectMany(extension => extension.Elements<P188.CommentRelationship>())
                .Any(reference => string.Equals(reference.Id?.Value, relationshipId, StringComparison.Ordinal))) {
                return;
            }
            P.SlideExtension? target = commentExtensions.FirstOrDefault(extension => !extension.HasChildren);
            if (target == null) {
                target = new P.SlideExtension { Uri = ModernCommentExtensionUri };
                extensions.Append(target);
            }
            target.Append(new P188.CommentRelationship { Id = relationshipId });
        }

        internal static void RemoveModernCommentRelationship(SlidePart slide, PowerPointCommentPart part) {
            string relationshipId = slide.GetIdOfPart(part);
            P.SlideExtensionList? extensions = slide.Slide?.SlideExtensionList;
            if (extensions == null) return;
            foreach (P.SlideExtension extension in extensions.Elements<P.SlideExtension>()
                .Where(IsModernCommentExtension).ToArray()) {
                P188.CommentRelationship[] references = extension.Elements<P188.CommentRelationship>()
                    .Where(reference => string.Equals(reference.Id?.Value, relationshipId, StringComparison.Ordinal)).ToArray();
                if (references.Length == 0) continue;
                foreach (P188.CommentRelationship reference in references) {
                    reference.Remove();
                }
                if (!extension.HasChildren && !extension.ExtendedAttributes.Any() && extension.MCAttributes == null) {
                    extension.Remove();
                }
            }
            if (!extensions.HasChildren && !extensions.ExtendedAttributes.Any() && extensions.MCAttributes == null) {
                extensions.Remove();
            }
        }

        private static bool IsModernCommentExtension(P.SlideExtension extension) =>
            string.Equals(extension.Uri?.Value, ModernCommentExtensionUri, StringComparison.OrdinalIgnoreCase);
    }
}
