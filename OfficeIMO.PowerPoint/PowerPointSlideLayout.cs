using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     A slide layout: the placeholders, shapes and background that slides using it inherit, on top of its master.
    /// </summary>
    public sealed class PowerPointSlideLayout : PowerPointMasterSurface {
        internal PowerPointSlideLayout(PowerPointDocument document, SlideLayoutPart part) : base(document, part) {
            LayoutPart = part;
        }

        internal SlideLayoutPart LayoutPart { get; }

        internal SlideLayout Root => LayoutPart.SlideLayout ??= new SlideLayout();

        internal override CommonSlideData CommonSlideData => Root.CommonSlideData ??= new CommonSlideData(new ShapeTree());

        /// <summary>The master that owns this layout.</summary>
        public PowerPointSlideMaster Master {
            get {
                ThrowIfRemoved();
                return Document.GetMaster(LayoutPart.SlideMasterPart
                    ?? throw new InvalidOperationException("The slide layout is not attached to a slide master."));
            }
        }

        /// <summary>Zero-based position of the layout within its master.</summary>
        public int Index => Master.OrderedLayoutParts().ToList().IndexOf(LayoutPart);

        /// <summary>Native layout type, or <see langword="null" /> when the layout does not declare one.</summary>
        public PowerPointSlideLayoutType? LayoutType {
            get {
                ThrowIfRemoved();
                return Root.Type?.Value.ToOfficeIMO();
            }
        }

        /// <summary>Whether slides using this layout also show the master's shapes.</summary>
        public bool ShowMasterShapes {
            get {
                ThrowIfRemoved();
                return Root.ShowMasterShapes?.Value != false;
            }
            set {
                ThrowIfRemoved();
                Root.ShowMasterShapes = value ? null : false;
            }
        }

        /// <summary>Slides that use this layout, in presentation order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides {
            get {
                ThrowIfRemoved();
                return Document.Slides.Where(slide => ReferenceEquals(slide.SlidePart.SlideLayoutPart, LayoutPart)).ToList();
            }
        }

        /// <summary>
        ///     Copies this layout, with its shapes, placeholders and pictures, into a new layout of the same master.
        /// </summary>
        public PowerPointSlideLayout Duplicate(string? name = null) {
            ThrowIfRemoved();
            var copy = (SlideLayout)Root.CloneNode(true);
            copy.Preserve = true;
            copy.UserDrawn = true;
            PowerPointSlideLayout duplicate = Document.GetLayout(Master.AttachLayout(copy, LayoutPart));
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

            foreach (PowerPointSlide slide in Slides) {
                slide.SetLayout(replacement);
            }

            Document.RemoveLayout(this);
        }
    }
}
