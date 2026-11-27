using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     A slide master: the theme, text styles, shapes and background shared by its layouts and their slides.
    /// </summary>
    public sealed class PowerPointSlideMaster : PowerPointMasterSurface {
        internal PowerPointSlideMaster(PowerPointDocument document, PowerPointPartRef part) : base(document, part) {
        }

        private protected override OfficeXmlChildOrder RootOrder => PSlideMasterView.Order;

        /// <summary>Zero-based position of the master in the presentation.</summary>
        public int Index {
            get {
                ThrowIfRemoved();
                return Document.MasterPartRefs().IndexOf(PartRef);
            }
        }

        /// <summary>Layouts owned by this master; positions match the layout indexes used by <see cref="PowerPointSlide.SetLayout(int, int)" />.</summary>
        public IReadOnlyList<PowerPointSlideLayout> Layouts {
            get {
                ThrowIfRemoved();
                return OrderedLayoutParts().Select(Document.GetLayout).ToList();
            }
        }

        /// <summary>Returns the first layout with the given name, or <see langword="null" />.</summary>
        public PowerPointSlideLayout? GetLayout(string name, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(name);
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return Layouts.FirstOrDefault(layout => string.Equals(layout.Name, name, comparison));
        }

        /// <summary>
        ///     Adds an empty custom layout at the end of this master's layout list. It shows the master's shapes and
        ///     background; add placeholders and shapes to design it.
        /// </summary>
        public PowerPointSlideLayout AddLayout(string name, PowerPointSlideLayoutType layoutType = PowerPointSlideLayoutType.Custom) {
            ThrowIfRemoved();
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            return Document.AddLayout(this, name, layoutType);
        }

        /// <summary>
        ///     Copies this master, its theme and all its layouts into a new master at the end of the presentation.
        /// </summary>
        public PowerPointSlideMaster Duplicate(string? name = null) {
            ThrowIfRemoved();
            PowerPointSlideMaster copy = Document.CloneMaster(this);
            copy.Name = name ?? (string.IsNullOrEmpty(Name) ? "Slide Master" : Name) + " Copy";
            return copy;
        }

        /// <summary>
        ///     Removes the master and its layouts. The presentation keeps at least one master, and a master whose
        ///     layouts are still used by slides cannot be removed; move those slides to another layout first.
        /// </summary>
        public void Remove() {
            ThrowIfRemoved();
            Document.RemoveMaster(this);
        }

        // Layout indexes follow the master's layout relationships, as the index-based layout APIs do.
        internal IEnumerable<PowerPointPartRef> OrderedLayoutParts() => PartRef.RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType);
    }
}
