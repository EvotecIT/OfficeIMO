using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private OfficeXmlNode _cachedSlideId;

        /// <summary>
        ///     Gets or sets whether the slide is hidden in slide show mode.
        /// </summary>
        public bool Hidden {
            get {
                bool? show = PartRef.Root.BooleanAttribute("show");
                return show.HasValue ? !show.Value : IsHiddenShowValue(GetSlideId().Attribute("show"));
            }
            set {
                OfficeXmlNode root = PartRef.Root;
                OfficeXmlNode id = GetSlideId();
                id.RemoveAttribute("show");
                root.SetBooleanAttribute("show", value ? false : null);
                PartRef.MarkEdited();
                OwnerDocument?.PresentationPartRef.MarkEdited();
            }
        }

        /// <summary>Hides the slide in slide show mode.</summary>
        public void Hide() => Hidden = true;

        /// <summary>Shows the slide in slide show mode.</summary>
        public void Show() => Hidden = false;

        private OfficeXmlNode GetSlideId() {
            const int presentationNamespace = OfficeXmlNamespaces.PresentationML;
            PowerPointPartRef slide = PartRef;
            PowerPointDocument document = OwnerDocument ?? throw new InvalidOperationException("Slide is not attached to a presentation.");
            if (!slide.IsSdk && document.GetPackageSlideEntry(slide.Name) is { } packageEntry) {
                return new OfficeXmlNode(packageEntry.Tree, packageEntry.Node);
            }

            PowerPointPartRef presentation = document.PresentationPartRef;
            OfficeXmlNode ids = presentation.Root.Child(presentationNamespace, "sldIdLst");
            if (ids.IsNone) throw new InvalidOperationException("Presentation has no slide list.");
            if (_cachedSlideId.Parent == ids && IdReferencesSlide(_cachedSlideId, presentation, slide)) {
                return _cachedSlideId;
            }

            foreach (OfficeXmlNode entry in ids.Children(presentationNamespace, "sldId")) {
                if (!IdReferencesSlide(entry, presentation, slide)) continue;
                return _cachedSlideId = entry;
            }

            throw new InvalidOperationException("Slide not found in presentation.");
        }

        private static bool IdReferencesSlide(OfficeXmlNode id, PowerPointPartRef presentation, PowerPointPartRef slide) =>
            id.Attribute(OfficeXmlNamespaces.Relationships, "id") is { } relationshipId
            && presentation.TryGetRelationship(relationshipId, out PowerPointRelationship relationship)
            && relationship.TargetPart.Equals(slide);

        /// <summary>Records the slide list entry the presentation just created or read for this slide.</summary>
        internal void RememberSlideId(OfficeXmlNode slideId) => _cachedSlideId = slideId;

        private static bool IsHiddenShowValue(string? showValue) =>
            string.Equals(showValue, "0", StringComparison.Ordinal) ||
            string.Equals(showValue, "false", StringComparison.OrdinalIgnoreCase);

        private void NormalizeHiddenSlideMarkup() {
            OfficeXmlNode root = PartRef.Root;
            OfficeXmlNode slideId = GetSlideId();
            string? legacyShowValue = slideId.Attribute("show");
            if (legacyShowValue == null) return;
            if (root.BooleanAttribute("show") == null && IsHiddenShowValue(legacyShowValue)) {
                root.SetBooleanAttribute("show", false);
                PartRef.MarkEdited();
            }

            slideId.RemoveAttribute("show");
            OwnerDocument?.PresentationPartRef.MarkEdited();
        }
    }
}
