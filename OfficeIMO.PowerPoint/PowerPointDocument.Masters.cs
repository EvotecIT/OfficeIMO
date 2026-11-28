using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        // Master and layout wrappers by part name, so each part has one. Created on first use: an initializer would run
        // in the constructors, which the SDK-bound members share.
        private Dictionary<string, PowerPointMasterSurface>? _masterSurfaceMap;

        private Dictionary<string, PowerPointMasterSurface> MasterSurfaces => _masterSurfaceMap ??= new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Slide masters in the presentation; positions match the master indexes used by the layout APIs.</summary>
        public IReadOnlyList<PowerPointSlideMaster> SlideMasters {
            get {
                ThrowIfDisposed();
                return MasterPartRefs().Select(GetMaster).ToList();
            }
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

        /// <summary>The slide master parts in relationship order, as the index-based APIs count them.</summary>
        internal List<PowerPointPartRef> MasterPartRefs() => PresentationPartRef.RelatedParts(PowerPointPartRef.SlideMasterRelationshipType).ToList();

        internal PowerPointSlideMaster GetMaster(PowerPointPartRef part) {
            if (!MasterSurfaces.TryGetValue(part.Name, out PowerPointMasterSurface? surface)) {
                surface = new PowerPointSlideMaster(this, part);
                MasterSurfaces[part.Name] = surface;
            }

            return (PowerPointSlideMaster)surface;
        }

        internal PowerPointSlideLayout GetLayout(PowerPointPartRef part) {
            if (!MasterSurfaces.TryGetValue(part.Name, out PowerPointMasterSurface? surface)) {
                surface = new PowerPointSlideLayout(this, part);
                MasterSurfaces[part.Name] = surface;
            }

            return (PowerPointSlideLayout)surface;
        }

        // Forgets a removed master or layout, so a later part with its name gets a new wrapper.
        private void ForgetSurface(PowerPointMasterSurface surface) {
            MasterSurfaces.Remove(surface.PartName);
            surface.MarkRemoved();
        }
    }
}
