using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     A slide layout: the placeholders, shapes and background that slides using it inherit.
    /// </summary>
    public sealed class PowerPointSlideLayout : PowerPointMasterSurface {
        internal PowerPointSlideLayout(PowerPointDocument document, PowerPointPartRef part) : base(document, part) {
        }

        private protected override OfficeXmlChildOrder RootOrder => PSlideLayoutView.Order;

        /// <summary>The master that owns this layout.</summary>
        public PowerPointSlideMaster Master {
            get {
                ThrowIfRemoved();
                PowerPointPartRef master = PartRef.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
                if (master.IsNone) throw new InvalidOperationException("The slide layout is not attached to a slide master.");
                return Document.GetMaster(master);
            }
        }

        /// <summary>Zero-based position of the layout within its master.</summary>
        public int Index => Master.OrderedLayoutParts().ToList().IndexOf(PartRef);

        /// <summary>Native layout type, or <see langword="null" /> when the layout does not declare one.</summary>
        public PowerPointSlideLayoutType? LayoutType {
            get {
                ThrowIfRemoved();
                string? token = Root.Attribute("type");
                if (token == null) return null;
                return PowerPointSlideLayoutTypeTokens.TryParse(token, out PowerPointSlideLayoutType type)
                    ? type
                    : throw new ArgumentOutOfRangeException(nameof(LayoutType), token, "Unsupported Open XML slide layout type.");
            }
        }

        /// <summary>Whether slides using this layout also show the master's shapes.</summary>
        public bool ShowMasterShapes {
            get {
                ThrowIfRemoved();
                return Root.BooleanAttribute("showMasterSp") != false;
            }
            set {
                ThrowIfRemoved();
                Root.SetAttribute("showMasterSp", value ? null : "0");
                PartRef.MarkEdited();
            }
        }

        /// <summary>Slides that use this layout, in presentation order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides {
            get {
                ThrowIfRemoved();
                PowerPointPartRef layout = PartRef;
                return Document.Slides.Where(slide => slide.PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType).Equals(layout)).ToList();
            }
        }

        /// <summary>
        ///     Copies this layout, with its shapes, placeholders and pictures, into a new layout of the same master.
        /// </summary>
        public PowerPointSlideLayout Duplicate(string? name = null) {
            ThrowIfRemoved();
            PowerPointSlideLayout duplicate = Document.DuplicateLayout(this);
            duplicate.Name = name ?? (string.IsNullOrEmpty(Name) ? "Layout" : Name) + " Copy";
            return duplicate;
        }

        /// <summary>
        ///     Removes the layout. It must not be used by any slide; use <see cref="Remove(PowerPointSlideLayout)" />
        ///     to move those slides first. A master keeps at least one layout.
        /// </summary>
        public void Remove() {
            ThrowIfRemoved();
            int used = Slides.Count;
            if (used > 0) {
                throw new InvalidOperationException("Layout '" + Name + "' is used by " + used
                    + " slide(s). Remove it with a replacement layout to move those slides.");
            }

            Document.RemoveLayout(this);
        }

        /// <summary>
        ///     Moves every slide using this layout to <paramref name="replacement" /> and then removes this layout.
        ///     Slide content is kept; placeholders are matched to the replacement's by type and index.
        /// </summary>
        public void Remove(PowerPointSlideLayout replacement) {
            ThrowIfRemoved();
            ArgumentNullException.ThrowIfNull(replacement);
            replacement.ThrowIfRemoved();
            if (ReferenceEquals(replacement, this)) {
                throw new ArgumentException("The replacement layout must be a different layout.", nameof(replacement));
            }
            if (!ReferenceEquals(replacement.Document, Document)) {
                throw new ArgumentException("The replacement layout must belong to the same presentation.", nameof(replacement));
            }

            Document.EnsureLayoutCanBeRemoved(this);
            foreach (PowerPointSlide slide in Slides) {
                slide.SetLayout(replacement);
            }

            Document.RemoveLayout(this);
        }
    }
}
