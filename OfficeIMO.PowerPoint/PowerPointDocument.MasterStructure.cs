using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Adding, copying and removing slide masters and layouts, which add and delete parts through the SDK. The SDK opens
    /// for these; reading and editing masters and layouts does not need it.
    /// </summary>
    public sealed partial class PowerPointDocument {
        // Slide master ids and slide layout ids share one identifier space that starts at 2^31.
        private const uint FirstMasterOrLayoutId = 2147483648U;

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
            PowerPointSlideMaster master = GetMaster(PowerPointPartRef.FromSdk(CloneSlideMasterPart(templateMaster, out _)));
            master.Name = name ?? "Slide Master " + SdkPresentationPart.SlideMasterParts.Count().ToString(System.Globalization.CultureInfo.InvariantCulture);
            return master;
        }

        internal PowerPointSlideMaster CloneMaster(PowerPointSlideMaster source) =>
            GetMaster(PowerPointPartRef.FromSdk(CloneSlideMasterPart(SdkMasterPart(source), out _)));

        internal void RemoveMaster(PowerPointSlideMaster master) {
            if (MasterPartRefs().Count <= 1) {
                throw new InvalidOperationException("A presentation must keep at least one slide master.");
            }

            int used = master.Layouts.Sum(layout => layout.Slides.Count);
            if (used > 0) {
                throw new InvalidOperationException("Slide master '" + master.Name + "' has layouts used by " + used
                    + " slide(s). Move those slides to another master's layout first.");
            }

            SlideMasterPart masterPart = SdkMasterPart(master);
            string relationshipId = SdkPresentationPart.GetIdOfPart(masterPart);
            PresentationRoot.SlideMasterIdList?.Elements<SlideMasterId>()
                .FirstOrDefault(id => PowerPointUtils.GetRelationshipIdValue(id) == relationshipId)?.Remove();
            foreach (PowerPointSlideLayout layout in master.Layouts) ForgetSurface(layout);
            ForgetSurface(master);
            SdkPresentationPart.DeletePart(masterPart);
        }

        internal void EnsureLayoutCanBeRemoved(PowerPointSlideLayout layout) {
            PowerPointSlideMaster master = layout.Master;
            if (master.OrderedLayoutParts().Count() <= 1) {
                throw new InvalidOperationException("Slide master '" + master.Name + "' must keep at least one layout.");
            }
        }

        internal void RemoveLayout(PowerPointSlideLayout layout) {
            EnsureLayoutCanBeRemoved(layout);
            PowerPointSlideMaster master = layout.Master;
            SlideMasterPart masterPart = SdkMasterPart(master);
            SlideLayoutPart layoutPart = (SlideLayoutPart)SdkPartOf(layout);

            string relationshipId = masterPart.GetIdOfPart(layoutPart);
            masterPart.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>()
                .FirstOrDefault(id => id.RelationshipId?.Value == relationshipId)?.Remove();
            ForgetSurface(layout);
            masterPart.DeletePart(layoutPart);
        }

        internal PowerPointSlideLayout AddLayout(PowerPointSlideMaster master, string name, PowerPointSlideLayoutType layoutType) {
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
            return GetLayout(PowerPointPartRef.FromSdk(AttachLayout(SdkMasterPart(master), layout, null)));
        }

        internal PowerPointSlideLayout DuplicateLayout(PowerPointSlideLayout source) {
            var sourcePart = (SlideLayoutPart)SdkPartOf(source);
            var copy = (SlideLayout)(sourcePart.SlideLayout ??= new SlideLayout()).CloneNode(true);
            copy.Preserve = true;
            copy.UserDrawn = true;
            return GetLayout(PowerPointPartRef.FromSdk(AttachLayout(SdkMasterPart(source.Master), copy, sourcePart)));
        }

        private SlideLayoutPart AttachLayout(SlideMasterPart masterPart, SlideLayout layout, SlideLayoutPart? source) {
            string relationshipId = GetNextRelationshipId(masterPart);
            SlideLayoutPart layoutPart;
            if (source != null) {
                layoutPart = CloneSlideLayoutPart(source, masterPart, relationshipId);
                layoutPart.SlideLayout = layout;
            } else {
                layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
                layoutPart.SlideLayout = layout;
                layoutPart.AddPart(masterPart);
            }

            SlideMaster root = masterPart.SlideMaster ??= new SlideMaster();
            SlideLayoutIdList list = root.SlideLayoutIdList ??= new SlideLayoutIdList();
            list.Append(new SlideLayoutId { Id = NextMasterOrLayoutId(), RelationshipId = relationshipId });
            return layoutPart;
        }

        internal uint NextMasterOrLayoutId() {
            uint max = FirstMasterOrLayoutId - 1U;
            foreach (SlideMasterId id in PresentationRoot.SlideMasterIdList?.Elements<SlideMasterId>()
                         ?? Enumerable.Empty<SlideMasterId>()) {
                if (id.Id?.Value > max) max = id.Id.Value;
            }

            foreach (SlideMasterPart masterPart in SdkPresentationPart.SlideMasterParts) {
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

        /// <summary>The shapes of a master or layout whose part the SDK holds or that has a chart, read through the SDK.</summary>
        internal IReadOnlyList<PowerPointShape> CreateSdkSurfaceShapes(PowerPointMasterSurface surface) {
            OpenXmlPart part = SdkPartOf(surface);
            var shapes = new List<PowerPointShape>();
            foreach (OfficeXmlNode element in surface.ShapeTree.Children()) {
                if (element.SdkElement is OpenXmlElement sdk && PowerPointSlide.CreateShapeFromElement(sdk, part) is { } shape) shapes.Add(shape);
            }
            return shapes;
        }

        /// <summary>A text box wrapper for a placeholder shape of a master or layout the SDK holds.</summary>
        internal PowerPointTextBox CreateSdkSurfaceTextBox(PowerPointMasterSurface surface, OfficeXmlNode element) =>
            new PowerPointTextBox((Shape)element.SdkElement!, SdkPartOf(surface));

        /// <summary>The root of a master's or layout's empty part, added through the SDK as the model added it.</summary>
        internal OfficeXmlNode EnsureSdkSurfaceRoot(PowerPointMasterSurface surface) {
            OpenXmlPart part = SdkPartOf(surface);
            if (part is SlideLayoutPart layout) layout.SlideLayout ??= new SlideLayout();
            else if (part is SlideMasterPart master) master.SlideMaster ??= new SlideMaster();
            return surface.PartRef.Root;
        }

        /// <summary>The wrapper of an element just added to a master or layout the SDK holds.</summary>
        internal PowerPointShape CreateSdkSurfaceShape(PowerPointMasterSurface surface, OfficeXmlNode element) =>
            PowerPointSlide.CreateShapeFromElement((OpenXmlElement)element.SdkElement!, SdkPartOf(surface))
            ?? throw new InvalidOperationException("The added element is not a shape.");

        /// <summary>Adds an image part to a master or layout the SDK holds and returns its relationship id.</summary>
        internal string AddSdkSurfaceImage(PowerPointMasterSurface surface, byte[] data, OfficeImageFormat format) {
            OpenXmlPart part = SdkPartOf(surface);
            string partUri = PowerPointPartFactory.GetIndexedPartUri(part.OpenXmlPackage, "ppt/media", "image",
                PowerPointPartFactory.GetImageExtension(format), allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(part, format.ToPartTypeInfo().ContentType, partUri);
            using (var stream = new MemoryStream(data, writable: false)) imagePart.FeedData(stream);
            return part.GetIdOfPart(imagePart);
        }

        // The SDK part of a master or layout, opening the SDK when it is not open.
        private OpenXmlPart SdkPartOf(PowerPointMasterSurface surface) {
            _ = SdkPresentationPart;
            return surface.PartRef.SdkPart as OpenXmlPart ?? throw new InvalidOperationException("The slide master or layout is no longer part of the presentation.");
        }

        /// <summary>The SDK part of a layout, opening the SDK when it is not open.</summary>
        internal SlideLayoutPart SdkLayoutPart(PowerPointSlideLayout layout) => (SlideLayoutPart)SdkPartOf(layout);

        private SlideMasterPart SdkMasterPart(PowerPointSlideMaster master) => (SlideMasterPart)SdkPartOf(master);

        private SlideLayoutPart? FindMatchingLayout(SlideLayoutPart sourceLayoutPart) {
            SlideLayout? sourceLayout = sourceLayoutPart.SlideLayout;
            if (sourceLayout == null) {
                return null;
            }

            string? sourceName = sourceLayout.CommonSlideData?.Name?.Value;
            SlideLayoutValues? sourceType = sourceLayout.Type?.Value;

            foreach (SlideMasterPart masterPart in SdkPresentationPart.SlideMasterParts) {
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    SlideLayout? candidateLayout = layoutPart.SlideLayout;
                    if (candidateLayout == null) {
                        continue;
                    }

                    SlideLayoutValues? candidateType = candidateLayout.Type?.Value;
                    if (sourceType.HasValue && candidateType != sourceType) {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(sourceName)) {
                        string? candidateName = candidateLayout.CommonSlideData?.Name?.Value;
                        if (!string.Equals(sourceName, candidateName, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                    }

                    return layoutPart;
                }
            }

            return null;
        }

        internal SlideLayoutPart GetSlideLayoutPart(int masterIndex, int layoutIndex) {
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            ArgumentOutOfRangeException.ThrowIfNegative(layoutIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(layoutIndex, layouts.Length);
            return layouts[layoutIndex];
        }

    }
}
