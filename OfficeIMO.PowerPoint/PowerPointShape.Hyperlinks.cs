using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Sets an external or internal-slide click hyperlink on this shape.</summary>
        /// <param name="uri">Absolute hyperlink target or a valid <c>#slide-N</c> fragment.</param>
        /// <param name="tooltip">Optional screen tip.</param>
        public void SetHyperlink(Uri uri, string? tooltip = null) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri && OwnerSlide != null
                && PowerPointHyperlinkResolver.TryResolveSlideFragment(
                    OwnerSlide.SlidePart, uri, out SlidePart? targetSlidePart)) {
                SetInternalHyperlink(targetSlidePart!, tooltip);
                return;
            }
            if (!uri.IsAbsoluteUri) {
                throw new ArgumentException(
                    "Shape hyperlinks require an absolute URI or a valid #slide-N fragment.",
                    nameof(uri));
            }
            if (OwnerSlide == null) {
                throw new InvalidOperationException("Hyperlinks require a shape attached to a slide.");
            }
            if (GetNonVisualDrawingProperties(create: true) == null) {
                throw new NotSupportedException(
                    "This shape type does not expose non-visual drawing properties.");
            }

            SlidePart slidePart = OwnerSlide.SlidePart;
            HyperlinkRelationship relationship = slidePart.AddHyperlinkRelationship(uri, true);
            var replacement = new A.HyperlinkOnClick { Id = relationship.Id };
            if (!string.IsNullOrWhiteSpace(tooltip)) replacement.Tooltip = tooltip;
            ReplaceShapeHyperlink(replacement);
        }

        /// <summary>Sets an internal click hyperlink to another slide in the same presentation.</summary>
        /// <param name="targetSlide">Target slide.</param>
        /// <param name="tooltip">Optional screen tip.</param>
        public void SetHyperlink(PowerPointSlide targetSlide,
            string? tooltip = null) {
            if (targetSlide == null) throw new ArgumentNullException(nameof(targetSlide));
            SetInternalHyperlink(targetSlide.SlidePart, tooltip);
        }

        private void SetInternalHyperlink(SlidePart targetSlidePart,
            string? tooltip) {
            if (OwnerSlide == null) {
                throw new InvalidOperationException(
                    "Hyperlinks require a shape attached to a slide.");
            }
            if (GetNonVisualDrawingProperties(create: true) == null) {
                throw new NotSupportedException(
                    "This shape type does not expose non-visual drawing properties.");
            }

            SlidePart sourceSlidePart = OwnerSlide.SlidePart;
            PresentationPart? sourcePresentation = sourceSlidePart.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault();
            PresentationPart? targetPresentation = targetSlidePart.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault();
            if (sourcePresentation == null
                || !ReferenceEquals(sourcePresentation, targetPresentation)) {
                throw new ArgumentException(
                    "The hyperlink target must belong to the same presentation.",
                    "targetSlide");
            }

            if (!sourceSlidePart.Parts.Any(pair => ReferenceEquals(
                    pair.OpenXmlPart, targetSlidePart))) {
                sourceSlidePart.AddPart(targetSlidePart);
            }
            string relationshipId = sourceSlidePart.GetIdOfPart(targetSlidePart);
            var replacement = new A.HyperlinkOnClick {
                Id = relationshipId,
                Action = "ppaction://hlinksldjump"
            };
            if (!string.IsNullOrWhiteSpace(tooltip)) replacement.Tooltip = tooltip;
            ReplaceShapeHyperlink(replacement);
        }

        /// <summary>Removes the click hyperlink from this shape.</summary>
        public void ClearHyperlink() => ReplaceShapeHyperlink(replacement: null);

        private void ReplaceShapeHyperlink(A.HyperlinkOnClick? replacement) {
            NonVisualDrawingProperties? properties = GetNonVisualDrawingProperties(create: replacement != null);
            if (properties == null && replacement == null) return;
            if (properties == null) {
                throw new NotSupportedException(
                    "This shape type does not expose non-visual drawing properties.");
            }
            A.HyperlinkOnClick[] previous = properties.Elements<A.HyperlinkOnClick>().ToArray();
            if (replacement != null) {
                A.HyperlinkSound? sound = previous.SelectMany(link =>
                    link.Elements<A.HyperlinkSound>()).FirstOrDefault();
                if (sound != null) replacement.Append((A.HyperlinkSound)sound.CloneNode(true));
                bool? endSound = previous.Select(link => link.EndSound?.Value)
                    .FirstOrDefault(value => value.HasValue);
                if (endSound.HasValue) replacement.EndSound = endSound.Value;
            }
            string[] previousRelationshipIds = previous
                .Select(link => link.Id?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            string[] soundRelationshipIds = previous
                .SelectMany(link => link.Elements<A.HyperlinkSound>())
                .Select(sound => sound.Embed?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            properties.RemoveAllChildren<A.HyperlinkOnClick>();
            if (replacement != null) properties.Append(replacement);
            if (OwnerSlide == null) return;

            SlidePart slidePart = OwnerSlide.SlidePart;
            foreach (string relationshipId in previousRelationshipIds) {
                if (ReferencesRelationship(slidePart.Slide, relationshipId)) {
                    continue;
                }
                HyperlinkRelationship? relationship = slidePart.HyperlinkRelationships
                    .FirstOrDefault(candidate => string.Equals(candidate.Id, relationshipId,
                        StringComparison.Ordinal));
                if (relationship != null) slidePart.DeleteReferenceRelationship(relationship);
            }
            foreach (string relationshipId in soundRelationshipIds) {
                PowerPointEmbeddedSound.RemoveIfUnused(slidePart, relationshipId);
            }
        }

        private static bool ReferencesRelationship(
            OpenXmlPartRootElement? root, string relationshipId) => root != null
            && (root.GetAttributes().Any(attribute => string.Equals(
                    attribute.NamespaceUri,
                    PowerPointUtils.RelationshipIdNamespace,
                    StringComparison.Ordinal)
                && string.Equals(attribute.Value, relationshipId,
                    StringComparison.Ordinal))
                || root.Descendants().Any(element => element.GetAttributes()
                    .Any(attribute => string.Equals(attribute.NamespaceUri,
                            PowerPointUtils.RelationshipIdNamespace,
                            StringComparison.Ordinal)
                        && string.Equals(attribute.Value, relationshipId,
                            StringComparison.Ordinal))));
    }
}
