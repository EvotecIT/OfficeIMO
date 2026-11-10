using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        // Slide master ids and slide layout ids share one identifier space that starts at 2^31.
        private const uint FirstMasterOrLayoutId = 2147483648U;

        private readonly Dictionary<OpenXmlPart, PowerPointMasterSurface> _masterSurfaces = new();

        /// <summary>Slide masters in the presentation; positions match the master indexes used by the layout APIs.</summary>
        public IReadOnlyList<PowerPointSlideMaster> SlideMasters {
            get {
                ThrowIfDisposed();
                return SlideMasterParts.Select(GetMaster).ToList();
            }
        }

        /// <summary>
        ///     Adds a new slide master with the default Office theme and the standard set of layouts, at the end of
        ///     the master list.
        /// </summary>
        public PowerPointSlideMaster AddSlideMaster(string? name = null) {
            ThrowIfDisposed();
            using var source = new MemoryStream(DefaultPackage.Value, writable: false);
            using PresentationDocument template = PresentationDocument.Open(source, false);
            SlideMasterPart templateMaster = template.PresentationPart?.SlideMasterParts.FirstOrDefault()
                ?? throw new InvalidOperationException("The default presentation scaffold has no slide master.");
            PowerPointSlideMaster master = CloneMaster(templateMaster);
            master.Name = name ?? "Slide Master " + SlideMasterParts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return master;
        }

        /// <summary>Adds a slide that uses the given layout.</summary>
        public PowerPointSlide AddSlide(PowerPointSlideLayout layout) {
            ArgumentNullException.ThrowIfNull(layout);
            layout.ThrowIfRemoved();
            if (!ReferenceEquals(layout.Document, this)) {
                throw new ArgumentException("The layout belongs to a different presentation.", nameof(layout));
            }

            return AddSlide(layout.Master.Index, layout.Index);
        }

        internal List<SlideMasterPart> SlideMasterParts => _presentationPart.SlideMasterParts.ToList();

        internal PowerPointSlideMaster GetMaster(SlideMasterPart part) {
            if (!_masterSurfaces.TryGetValue(part, out PowerPointMasterSurface? surface)) {
                surface = new PowerPointSlideMaster(this, part);
                _masterSurfaces[part] = surface;
            }

            return (PowerPointSlideMaster)surface;
        }

        internal PowerPointSlideLayout GetLayout(SlideLayoutPart part) {
            if (!_masterSurfaces.TryGetValue(part, out PowerPointMasterSurface? surface)) {
                surface = new PowerPointSlideLayout(this, part);
                _masterSurfaces[part] = surface;
            }

            return (PowerPointSlideLayout)surface;
        }

        internal PowerPointSlideMaster CloneMaster(SlideMasterPart source) {
            SlideMasterPart clone = CloneSlideMasterPart(source, out _);
            return GetMaster(clone);
        }

        internal void RemoveMaster(PowerPointSlideMaster master) {
            List<SlideMasterPart> masters = SlideMasterParts;
            if (masters.Count <= 1) {
                throw new InvalidOperationException("A presentation must keep at least one slide master.");
            }

            int used = master.Layouts.Sum(layout => layout.Slides.Count);
            if (used > 0) {
                throw new InvalidOperationException("Slide master '" + master.Name + "' has layouts used by " + used
                    + " slide(s). Move those slides to another master's layout first.");
            }

            string relationshipId = _presentationPart.GetIdOfPart(master.MasterPart);
            PresentationRoot.SlideMasterIdList?.Elements<SlideMasterId>()
                .FirstOrDefault(id => PowerPointUtils.GetRelationshipIdValue(id) == relationshipId)?.Remove();
            foreach (PowerPointSlideLayout layout in master.Layouts) {
                _masterSurfaces.Remove(layout.LayoutPart);
                layout.MarkRemoved();
            }

            _presentationPart.DeletePart(master.MasterPart);
            _masterSurfaces.Remove(master.MasterPart);
            master.MarkRemoved();
        }

        internal void RemoveLayout(PowerPointSlideLayout layout) {
            PowerPointSlideMaster master = layout.Master;
            if (master.MasterPart.SlideLayoutParts.Count() <= 1) {
                throw new InvalidOperationException("Slide master '" + master.Name + "' must keep at least one layout.");
            }

            string relationshipId = master.MasterPart.GetIdOfPart(layout.LayoutPart);
            master.Root.SlideLayoutIdList?.Elements<SlideLayoutId>()
                .FirstOrDefault(id => id.RelationshipId?.Value == relationshipId)?.Remove();
            master.MasterPart.DeletePart(layout.LayoutPart);
            _masterSurfaces.Remove(layout.LayoutPart);
            layout.MarkRemoved();
        }

        internal uint NextMasterOrLayoutId() {
            uint max = FirstMasterOrLayoutId - 1U;
            foreach (SlideMasterId id in PresentationRoot.SlideMasterIdList?.Elements<SlideMasterId>()
                         ?? Enumerable.Empty<SlideMasterId>()) {
                if (id.Id?.Value > max) max = id.Id.Value;
            }

            foreach (SlideMasterPart masterPart in _presentationPart.SlideMasterParts) {
                foreach (SlideLayoutId id in masterPart.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>()
                             ?? Enumerable.Empty<SlideLayoutId>()) {
                    if (id.Id?.Value > max) max = id.Id.Value;
                }
            }

            if (max == uint.MaxValue) {
                throw new InvalidOperationException("The presentation has exhausted the slide master and layout identifier range.");
            }

            return max + 1U;
        }

        internal static string NextRelationshipId(OpenXmlPartContainer container) => GetNextRelationshipId(container);

        internal static SlideLayoutPart CloneLayoutPart(SlideLayoutPart source, SlideMasterPart master, string relationshipId) =>
            CloneSlideLayoutPart(source, master, relationshipId);
    }
}
