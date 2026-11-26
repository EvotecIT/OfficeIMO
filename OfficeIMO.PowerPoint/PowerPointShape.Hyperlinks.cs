using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Sets an external or internal-slide click hyperlink on this shape.</summary>
        /// <param name="uri">Absolute hyperlink target or a valid <c>#slide-N</c> fragment.</param>
        /// <param name="tooltip">Optional screen tip.</param>
        public void SetHyperlink(Uri uri, string? tooltip = null) {
            ArgumentNullException.ThrowIfNull(uri);
            if (!uri.IsAbsoluteUri && OwnerSlide?.OwnerDocument is { } document
                && PowerPointHyperlinkResolver.ResolveSlideFragment(document, uri) is { IsNone: false } target) {
                SetInternalHyperlink(target, tooltip);
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
            RequireNonVisualDrawing();

            string relationshipId = HyperlinkOwnerPart.AddExternalHyperlink(uri);
            ReplaceShapeHyperlink(relationshipId, action: null, tooltip, remove: false);
        }

        /// <summary>Sets an internal click hyperlink to another slide in the same presentation.</summary>
        /// <param name="targetSlide">Target slide.</param>
        /// <param name="tooltip">Optional screen tip.</param>
        public void SetHyperlink(PowerPointSlide targetSlide,
            string? tooltip = null) {
            ArgumentNullException.ThrowIfNull(targetSlide);
            SetInternalHyperlink(targetSlide.PartRef, tooltip);
        }

        // The part the shape's hyperlink relationships belong to: the part that holds it, or its slide's part.
        private PowerPointPartRef HyperlinkOwnerPart {
            get {
                PowerPointPartRef owner = OwnerPartRef;
                return owner.IsNone ? OwnerSlide!.PartRef : owner;
            }
        }

        private void SetInternalHyperlink(PowerPointPartRef targetSlidePart, string? tooltip) {
            if (OwnerSlide == null) {
                throw new InvalidOperationException(
                    "Hyperlinks require a shape attached to a slide.");
            }
            RequireNonVisualDrawing();
            if (OwnerSlide.OwnerDocument is null || !ReferenceEquals(OwnerSlide.OwnerDocument, targetSlidePart.Document)) {
                throw new ArgumentException(
                    "The hyperlink target must belong to the same presentation.",
                    "targetSlide");
            }

            PowerPointPartRef owner = HyperlinkOwnerPart;
            string relationshipId = owner.Kind != PowerPointPartKind.Slide
                ? owner.GetOrAddRelativeHyperlink(targetSlidePart)
                : owner.GetOrAddSlideRelationship(targetSlidePart);
            ReplaceShapeHyperlink(relationshipId, PowerPointHyperlinkResolver.SlideJumpAction, tooltip, remove: false);
        }

        /// <summary>Removes the click hyperlink from this shape.</summary>
        public void ClearHyperlink() => ReplaceShapeHyperlink(relationshipId: null, action: null, tooltip: null, remove: true);

        // Replaces the shape's click hyperlinks with one to the relationship, keeping the previous one's sound, and
        // removes relationships nothing uses any more.
        private void ReplaceShapeHyperlink(string? relationshipId, string? action, string? tooltip, bool remove) {
            OfficeXmlNode properties = NonVisualDrawingElement(create: !remove);
            if (properties.IsNone && remove) return;
            if (properties.IsNone) {
                throw new NotSupportedException(
                    "This shape type does not expose non-visual drawing properties.");
            }
            List<OfficeXmlNode> previous = properties.Children(OfficeXmlNamespaces.DrawingML, "hlinkClick").ToList();
            OfficeXmlNode sound = default;
            bool? endSound = null;
            var relationshipIds = new List<string>();
            var soundIds = new List<string>();
            foreach (OfficeXmlNode link in previous) {
                foreach (OfficeXmlNode linkSound in link.Children(OfficeXmlNamespaces.DrawingML, "snd")) {
                    if (sound.IsNone) sound = linkSound.Clone();
                    if (linkSound.Attribute(OfficeXmlNamespaces.Relationships, "embed") is { Length: > 0 } embed && !soundIds.Contains(embed)) soundIds.Add(embed);
                }
                endSound ??= link.BooleanAttribute("endSnd");
                if (link.Attribute(OfficeXmlNamespaces.Relationships, "id") is { Length: > 0 } id && !relationshipIds.Contains(id)) relationshipIds.Add(id);
            }
            foreach (OfficeXmlNode link in previous) link.Remove();
            if (!remove) {
                OfficeXmlNode replacement = properties.CreateElement(OfficeXmlNamespaces.DrawingML, "hlinkClick");
                replacement.SetAttribute(OfficeXmlNamespaces.Relationships, "id", relationshipId);
                if (action != null) replacement.SetAttribute("action", action);
                if (!string.IsNullOrWhiteSpace(tooltip)) replacement.SetAttribute("tooltip", tooltip);
                if (endSound.HasValue) replacement.SetBooleanAttribute("endSnd", endSound.Value);
                replacement = properties.Append(replacement);
                if (!sound.IsNone) replacement.Append(sound);
            }
            OwnerPartRef.MarkEdited();
            if (OwnerSlide == null) return;

            PowerPointPartRef owner = HyperlinkOwnerPart;
            foreach (string id in relationshipIds) owner.RemoveRelationshipIfUnused(id, default);
            foreach (string id in soundIds) owner.RemoveRelationshipIfUnused(id, default);
        }
    }
}
