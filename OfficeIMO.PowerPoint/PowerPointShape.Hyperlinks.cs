using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Sets an external click hyperlink on this shape.</summary>
        /// <param name="uri">Absolute hyperlink target.</param>
        /// <param name="tooltip">Optional screen tip.</param>
        public void SetHyperlink(Uri uri, string? tooltip = null) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri) {
                throw new ArgumentException("Shape hyperlinks require an absolute URI.", nameof(uri));
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
                if (slidePart.Slide?.Descendants<A.HyperlinkOnClick>()
                        .Any(link => string.Equals(link.Id?.Value, relationshipId,
                            StringComparison.Ordinal)) == true) {
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
    }
}
