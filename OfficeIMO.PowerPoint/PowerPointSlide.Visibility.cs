using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private SlideId? _cachedSlideId;

        /// <summary>
        ///     Gets or sets whether the slide is hidden in slide show mode.
        /// </summary>
        public bool Hidden {
            get {
                if (SlideRoot.Show?.Value != null) {
                    return !SlideRoot.Show.Value;
                }

                return IsHiddenShowValue(GetLegacySlideIdShowValue(GetSlideId()));
            }
            set {
                SlideId slideId = GetSlideId();
                slideId.RemoveAttribute("show", string.Empty);

                if (value) {
                    SlideRoot.Show = false;
                } else {
                    SlideRoot.Show = null;
                }
            }
        }

        /// <summary>
        ///     Hides the slide in slide show mode.
        /// </summary>
        public void Hide() => Hidden = true;

        /// <summary>
        ///     Shows the slide in slide show mode.
        /// </summary>
        public void Show() => Hidden = false;

        private SlideId GetSlideId() {
            PresentationPart presentationPart = _slidePart.GetParentPresentationPart()
                ?? throw new InvalidOperationException("Slide is not attached to a presentation.");

            SlideIdList? slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slide list.");
            }

            // The slide list entry is found once and reused while it still points at this slide part.
            SlideId? cached = _cachedSlideId;
            if (cached != null && ReferenceEquals(cached.Parent, slideIdList)) {
                string? cachedRelId = cached.RelationshipId?.Value;
                if (cachedRelId is { Length: > 0 }
                    && presentationPart.TryGetPartById(cachedRelId, out OpenXmlPart? cachedPart)
                    && ReferenceEquals(cachedPart, _slidePart)) {
                    return cached;
                }
            }

            string relId = presentationPart.GetIdOfPart(_slidePart);
            SlideId? slideId = null;
            foreach (SlideId candidate in slideIdList.Elements<SlideId>()) {
                if (candidate.RelationshipId?.Value == relId) {
                    slideId = candidate;
                    break;
                }
            }

            if (slideId == null) {
                throw new InvalidOperationException("Slide not found in presentation.");
            }

            _cachedSlideId = slideId;
            return slideId;
        }

        /// <summary>Records the slide list entry the presentation just created or read for this slide.</summary>
        internal void RememberSlideId(SlideId slideId) => _cachedSlideId = slideId;

        private static string? GetLegacySlideIdShowValue(SlideId slideId) {
            // The legacy show flag is not part of the sldId schema, so it is only ever an extended attribute.
            if (!slideId.HasAttributes) return null;
            foreach (OpenXmlAttribute attribute in slideId.ExtendedAttributes) {
                if (attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return attribute.Value;
                }
            }
            return null;
        }

        private static bool IsHiddenShowValue(string? showValue) {
            if (string.IsNullOrEmpty(showValue)) {
                return false;
            }

            return string.Equals(showValue, "0", StringComparison.Ordinal) ||
                   string.Equals(showValue, "false", StringComparison.OrdinalIgnoreCase);
        }

        private void NormalizeHiddenSlideMarkup() {
            SlideId slideId = GetSlideId();
            string? legacyShowValue = GetLegacySlideIdShowValue(slideId);
            if (SlideRoot.Show?.Value == null && IsHiddenShowValue(legacyShowValue)) {
                SlideRoot.Show = false;
            }

            slideId.RemoveAttribute("show", string.Empty);
        }
    }
}
