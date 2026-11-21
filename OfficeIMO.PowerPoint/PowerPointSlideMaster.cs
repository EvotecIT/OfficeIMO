using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     A slide master: the shapes, placeholders, background and theme shared by its layouts and their slides.
    /// </summary>
    public sealed class PowerPointSlideMaster : PowerPointMasterSurface {
        internal PowerPointSlideMaster(PowerPointDocument document, SlideMasterPart part) : base(document, part) {
            MasterPart = part;
        }

        internal SlideMasterPart MasterPart { get; }

        internal SlideMaster Root => MasterPart.SlideMaster ??= new SlideMaster();

        internal override CommonSlideData CommonSlideData => Root.CommonSlideData ??= new CommonSlideData(new ShapeTree());

        /// <summary>Zero-based position of the master in the presentation.</summary>
        public int Index {
            get {
                ThrowIfRemoved();
                return Document.SlideMasterParts.IndexOf(MasterPart);
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
            var layout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    PowerPointUtils.CreateDefaultGroupShapeProperties())) { Name = name },
                new ColorMapOverride(new A.MasterColorMapping())) {
                Type = layoutType.ToOpenXml(),
                Preserve = true,
                UserDrawn = true
            };
            return Document.GetLayout(AttachLayout(layout));
        }

        /// <summary>
        ///     Copies this master, its theme and all its layouts into a new master at the end of the presentation.
        /// </summary>
        public PowerPointSlideMaster Duplicate(string? name = null) {
            ThrowIfRemoved();
            PowerPointSlideMaster copy = Document.CloneMaster(MasterPart);
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

        internal SlideLayoutPart AttachLayout(SlideLayout layout, SlideLayoutPart? source = null) {
            string relationshipId = PowerPointDocument.NextRelationshipId(MasterPart);
            SlideLayoutPart layoutPart;
            if (source != null) {
                layoutPart = PowerPointDocument.CloneLayoutPart(source, MasterPart, relationshipId);
                layoutPart.SlideLayout = layout;
            } else {
                layoutPart = MasterPart.AddNewPart<SlideLayoutPart>(relationshipId);
                layoutPart.SlideLayout = layout;
                layoutPart.AddPart(MasterPart);
            }

            SlideLayoutIdList list = Root.SlideLayoutIdList ??= new SlideLayoutIdList();
            list.Append(new SlideLayoutId { Id = Document.NextMasterOrLayoutId(), RelationshipId = relationshipId });
            return layoutPart;
        }

        // Layout indexes follow the master's layout relationships, as the index-based layout APIs do.
        internal IEnumerable<SlideLayoutPart> OrderedLayoutParts() => MasterPart.SlideLayoutParts;
    }
}
