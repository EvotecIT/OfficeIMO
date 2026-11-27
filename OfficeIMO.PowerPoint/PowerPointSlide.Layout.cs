using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Gets the index of the layout used by this slide.
        /// </summary>
        public int LayoutIndex {
            get {
                PowerPointPartRef layout = PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                if (layout.IsNone) {
                    return -1;
                }

                PowerPointPartRef master = layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
                return master.RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType).ToList().IndexOf(layout);
            }
        }

        /// <summary>
        ///     Sets the slide layout using master and layout indexes.
        /// </summary>
        public void SetLayout(int masterIndex, int layoutIndex) {
            PowerPointDocument document = OwnerDocument ?? throw new InvalidOperationException("The slide is not part of a presentation.");
            PowerPointPartRef layout = document.RequireLayoutPartRef(masterIndex, layoutIndex);
            PartRef.ReplaceRelatedPart(PowerPointPartRef.SlideLayoutRelationshipType, layout);
        }

        /// <summary>
        ///     The layout this slide uses, or <see langword="null" /> when the slide has no layout.
        /// </summary>
        public PowerPointSlideLayout? Layout {
            get {
                PowerPointPartRef layout = PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                return layout.IsNone || OwnerDocument == null ? null : OwnerDocument.GetLayout(layout);
            }
        }

        /// <summary>
        ///     Moves the slide to another layout of the same presentation. Slide content is kept; placeholders are
        ///     matched to the new layout's by type and index.
        /// </summary>
        public void SetLayout(PowerPointSlideLayout layout) {
            ArgumentNullException.ThrowIfNull(layout);
            layout.ThrowIfRemoved();
            if (OwnerDocument == null || !ReferenceEquals(layout.Document, OwnerDocument)) {
                throw new ArgumentException("The layout belongs to a different presentation.", nameof(layout));
            }

            PowerPointPartRef target = layout.PartRef;
            if (PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType).Equals(target)) return;
            PartRef.ReplaceRelatedPart(PowerPointPartRef.SlideLayoutRelationshipType, target);
        }

        /// <summary>
        ///     Sets the slide layout using a layout type.
        /// </summary>
        public void SetLayout(PowerPointSlideLayoutType layoutType, int masterIndex = 0) {
            PowerPointDocument document = OwnerDocument ?? throw new InvalidOperationException("The slide is not part of a presentation.");
            SetLayout(masterIndex, document.GetLayoutIndex(layoutType, masterIndex));
        }

        /// <summary>
        ///     Sets the slide layout using a layout name.
        /// </summary>
        public void SetLayout(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(layoutName);
            PowerPointDocument document = OwnerDocument ?? throw new InvalidOperationException("The slide is not part of a presentation.");
            SetLayout(masterIndex, document.GetLayoutIndex(layoutName, masterIndex, ignoreCase));
        }

        /// <summary>
        ///     Textboxes that map to placeholders in the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointTextBox> Placeholders =>
            TextBoxes.Where(tb => tb.IsPlaceholder).ToList();

        /// <summary>
        ///     Retrieves the first placeholder textbox matching the specified type.
        /// </summary>
        public PowerPointTextBox? GetPlaceholder(PowerPointPlaceholderType placeholderType, uint? index = null) {
            IEnumerable<PowerPointTextBox> matches = TextBoxes
                .Where(tb => tb.PlaceholderType == placeholderType);

            if (index != null) {
                matches = matches.Where(tb => tb.PlaceholderIndex == index);
            }

            return matches.FirstOrDefault();
        }

        /// <summary>
        ///     Retrieves placeholders defined by the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointLayoutPlaceholderInfo> GetLayoutPlaceholders() {
            OfficeXmlNode tree = PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType).Root
                .Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "spTree");
            List<PowerPointLayoutPlaceholderInfo> placeholders = new();
            foreach (OfficeXmlNode element in tree.Children()) {
                OfficeXmlNode placeholder = PlaceholderOf(element);
                if (placeholder.IsNone) {
                    continue;
                }

                string name = element.FirstChild.Child(OfficeXmlNamespaces.PresentationML, "cNvPr").Attribute("name") ?? string.Empty;
                PowerPointPlaceholderType? type = placeholder.Attribute("type") is { } token
                    && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderType parsed) ? parsed : null;
                placeholders.Add(new PowerPointLayoutPlaceholderInfo(name, type, placeholder.UInt32Attribute("idx"), ElementBounds(element)));
            }

            return placeholders;
        }

        /// <summary>
        ///     Retrieves a layout placeholder by type and optional index.
        /// </summary>
        public PowerPointLayoutPlaceholderInfo? GetLayoutPlaceholder(PowerPointPlaceholderType placeholderType, uint? index = null) {
            foreach (PowerPointLayoutPlaceholderInfo placeholder in GetLayoutPlaceholders()) {
                if (placeholder.PlaceholderType != placeholderType) {
                    continue;
                }

                if (index != null && placeholder.PlaceholderIndex != index) {
                    continue;
                }

                return placeholder;
            }

            return null;
        }

        /// <summary>
        ///     Retrieves layout placeholder bounds by type and optional index.
        /// </summary>
        public PowerPointLayoutBox? GetLayoutPlaceholderBounds(PowerPointPlaceholderType placeholderType, uint? index = null) {
            PowerPointLayoutPlaceholderInfo? placeholder = GetLayoutPlaceholder(placeholderType, index);
            return placeholder?.Bounds;
        }
    }
}
