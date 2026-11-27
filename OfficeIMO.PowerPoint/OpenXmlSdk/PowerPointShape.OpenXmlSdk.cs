using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a shape: its constructor over an SDK element, binding to the SDK element, and the SDK element
    /// helpers that shape members not yet on the element handle use.
    /// </summary>
    public abstract partial class PowerPointShape {
        private OpenXmlElement? _element;
        private OpenXmlPart? _relationshipOwnerPart;

        internal PowerPointShape(OpenXmlElement element, OpenXmlPartContainer? ownerPart = null) {
            _element = element;
            RelationshipOwnerPart = ownerPart as OpenXmlPart;
        }

        // A shape of a slide whose SDK part is not open: it reads and edits its element in the slide's package tree, and
        // anything it does not do there opens the SDK and binds the shape to the same element there.
        internal PowerPointShape(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            OwnerSlide = slide;
            PackageTree = tree;
            PackageNode = node;
        }

        internal OpenXmlElement Element => _element ?? BindSdkElement();

        /// <summary>Binds a package-backed shape to its SDK element now, while the slide the SDK opened is unchanged.</summary>
        internal void BindSdk() {
            if (_element != null || PackageTree == null) return;
            try {
                BindSdkElement();
            } catch (InvalidOperationException) {
                // A shape the opened slide doesn't have where expected reports it when it is used, not while the SDK opens.
            }
        }
        internal PowerPointSlide? OwnerSlide { get; private set; }

        internal OpenXmlPart? RelationshipOwnerPart {
            get => _relationshipOwnerPart ?? (PackageTree != null ? OwnerSlide?.SlidePart : null);
            private set => _relationshipOwnerPart = value;
        }

        /// <summary>The slide tree that holds the shape while it works without the SDK, and its element there.</summary>
        internal OfficeXmlTree? PackageTree { get; }
        internal int PackageNode { get; }

        /// <summary>Whether the shape works on its slide's package tree rather than an SDK element.</summary>
        internal bool IsPackageBacked => _element == null;

        private OpenXmlElement BindSdkElement() {
            OpenXmlElement element = PowerPointSlide.BindPackageElement(OwnerSlide!, PackageTree!, PackageNode);
            _element = element;
            return element;
        }

        internal PowerPointShape AttachTo(PowerPointSlide slide, OpenXmlPartContainer? ownerPart = null) {
            OwnerSlide = slide;
            RelationshipOwnerPart = ownerPart as OpenXmlPart ?? RelationshipOwnerPart ?? slide.SlidePart;
            return this;
        }

        private NonVisualDrawingProperties? GetNonVisualDrawingProperties(bool create) {
            switch (Element) {
                case Shape s: {
                    if (create) {
                        s.NonVisualShapeProperties ??= new NonVisualShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties());
                        s.NonVisualShapeProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return s.NonVisualShapeProperties?.NonVisualDrawingProperties;
                }
                case ConnectionShape c: {
                    if (create) {
                        c.NonVisualConnectionShapeProperties ??= new NonVisualConnectionShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualConnectorShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        c.NonVisualConnectionShapeProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties;
                }
                case Picture p: {
                    if (create) {
                        p.NonVisualPictureProperties ??= new NonVisualPictureProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        p.NonVisualPictureProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return p.NonVisualPictureProperties?.NonVisualDrawingProperties;
                }
                case GraphicFrame g: {
                    if (create) {
                        g.NonVisualGraphicFrameProperties ??= new NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGraphicFrameDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        g.NonVisualGraphicFrameProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                }
                case GroupShape g: {
                    if (create) {
                        g.NonVisualGroupShapeProperties ??= new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        g.NonVisualGroupShapeProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                }
                default:
                    return null;
            }
        }

        internal bool TryGetExportBoundsPoints(out double left, out double top, out double width, out double height) {
            bool found = TryGetBoundsPoints(out left, out top, out width, out height);
            if (found && width > 0D && height > 0D) return true;
            if (OwnerSlide != null && GetPlaceholderShape() is { } placeholder) {
                PowerPointLayoutBox? bounds = OwnerSlide.GetExportPlaceholderBounds(placeholder);
                if (bounds.HasValue) {
                    left = bounds.Value.Left.ToPoints();
                    top = bounds.Value.Top.ToPoints();
                    width = bounds.Value.Width.ToPoints();
                    height = bounds.Value.Height.ToPoints();
                    return width > 0D && height > 0D;
                }
            }
            return found;
        }

        internal bool HasPlaceholderShape => GetPlaceholderShape() != null;

        private PlaceholderShape? GetPlaceholderShape() {
            return Element switch {
                Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                ConnectionShape c => c.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                Picture p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
        }

        private static bool IsDrawingElement(OpenXmlElement element) =>
            element is Shape or ConnectionShape or Picture or GraphicFrame or GroupShape;
    }
}
