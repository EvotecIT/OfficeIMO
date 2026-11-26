using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Left of the shape.</summary>
        public OfficeLength Left {
            get => OfficeLength.Emus(TransformPoint("off", create: false).Int64Attribute("x") ?? 0L);
            set {
                SetTransformValue("off", "x", value.ToEmus());
                OnGeometryChanged();
            }
        }
        /// <summary>Top of the shape.</summary>
        public OfficeLength Top {
            get => OfficeLength.Emus(TransformPoint("off", create: false).Int64Attribute("y") ?? 0L);
            set {
                SetTransformValue("off", "y", value.ToEmus());
                OnGeometryChanged();
            }
        }
        /// <summary>Width of the shape.</summary>
        public OfficeLength Width {
            get => OfficeLength.Emus(TransformPoint("ext", create: false).Int64Attribute("cx") ?? 0L);
            set {
                SetTransformValue("ext", "cx", value.ToEmus());
                OnGeometryChanged();
            }
        }
        /// <summary>Height of the shape.</summary>
        public OfficeLength Height {
            get => OfficeLength.Emus(TransformPoint("ext", create: false).Int64Attribute("cy") ?? 0L);
            set {
                SetTransformValue("ext", "cy", value.ToEmus());
                OnGeometryChanged();
            }
        }

        /// <summary>
        ///     Connection sites where connectors can be glued, in slide coordinates. Empty when the geometry
        ///     declares no sites this library can resolve.
        /// </summary>
        public IReadOnlyList<PowerPointConnectionSite> ConnectionSites =>
            PowerPointConnectorGeometry.TryGetSites(this, out List<PowerPointConnectionSite> sites)
                ? sites
                : Array.Empty<PowerPointConnectionSite>();

        // Glued connectors follow geometry edits made through this API.
        internal void OnGeometryChanged() => OwnerSlide?.RerouteConnectorsAttachedTo(this);

        internal bool TryGetBoundsPoints(out double left, out double top, out double width, out double height) {
            left = 0D;
            top = 0D;
            width = 0D;
            height = 0D;
            if (!TryGetTransformValues(out long x, out long y, out long cx, out long cy)) {
                return false;
            }

            left = PowerPointUnits.ToPoints(x);
            top = PowerPointUnits.ToPoints(y);
            width = PowerPointUnits.ToPoints(cx);
            height = PowerPointUnits.ToPoints(cy);
            return true;
        }
        /// <summary>Right of the shape.</summary>
        public OfficeLength Right {
            get => Left + Width;
            set => Left = value - Width;
        }
        /// <summary>Bottom of the shape.</summary>
        public OfficeLength Bottom {
            get => Top + Height;
            set => Top = value - Height;
        }
        /// <summary>CenterX of the shape.</summary>
        public OfficeLength CenterX {
            get => OfficeLength.Emus(Left.ToEmus() + (long)Math.Round(Width.ToEmus() / 2d));
            set => Left = OfficeLength.Emus((long)Math.Round(value.ToEmus() - Width.ToEmus() / 2d));
        }
        /// <summary>CenterY of the shape.</summary>
        public OfficeLength CenterY {
            get => OfficeLength.Emus(Top.ToEmus() + (long)Math.Round(Height.ToEmus() / 2d));
            set => Top = OfficeLength.Emus((long)Math.Round(value.ToEmus() - Height.ToEmus() / 2d));
        }

        /// <summary>
        ///     Gets or sets the shape bounds.
        /// </summary>
        public PowerPointLayoutBox Bounds {
            get => new PowerPointLayoutBox(Left, Top, Width, Height);
            set {
                Left = value.Left;
                Top = value.Top;
                Width = value.Width;
                Height = value.Height;
            }
        }

        /// <summary>
        ///     Sets position with explicit lengths.
        /// </summary>
        public PowerPointShape SetPosition(OfficeLength left, OfficeLength top) {
            Left = left;
            Top = top;
            return this;
        }

        /// <summary>
        ///     Sets size with explicit lengths.
        /// </summary>
        public PowerPointShape SetSize(OfficeLength width, OfficeLength height) {
            Width = width;
            Height = height;
            return this;
        }

        /// <summary>
        ///     Resizes the shape while keeping the specified anchor fixed.
        /// </summary>
        public PowerPointShape Resize(OfficeLength width, OfficeLength height, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            long widthValue = width.ToEmus();
            long heightValue = height.ToEmus();
            ArgumentOutOfRangeException.ThrowIfNegative(widthValue, nameof(width));
            ArgumentOutOfRangeException.ThrowIfNegative(heightValue, nameof(height));

            PowerPointLayoutBox bounds = Bounds;
            (double anchorX, double anchorY) = ResolveAnchorPoint(bounds, anchor);
            (double offsetX, double offsetY) = ResolveAnchorOffset(widthValue, heightValue, anchor);

            Width = OfficeLength.Emus(widthValue);
            Height = OfficeLength.Emus(heightValue);
            Left = OfficeLength.Emus((long)Math.Round(anchorX - offsetX));
            Top = OfficeLength.Emus((long)Math.Round(anchorY - offsetY));
            return this;
        }

        /// <summary>
        ///     Scales the shape uniformly using the specified anchor.
        /// </summary>
#pragma warning disable RS0026 // Existing uniform and per-axis Scale overloads; only the owner return type changed.
        public PowerPointShape Scale(double scale, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.Center) {
#pragma warning restore RS0026
            if (scale < 0) {
                throw new ArgumentOutOfRangeException(nameof(scale));
            }

            long newWidth = (long)Math.Round(Width.ToEmus() * scale);
            long newHeight = (long)Math.Round(Height.ToEmus() * scale);
            Resize(OfficeLength.Emus(newWidth), OfficeLength.Emus(newHeight), anchor);
            return this;
        }

        /// <summary>
        ///     Scales the shape using separate X/Y factors and the specified anchor.
        /// </summary>
#pragma warning disable RS0026 // Existing uniform and per-axis Scale overloads; only the owner return type changed.
        public PowerPointShape Scale(double scaleX, double scaleY, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.Center) {
#pragma warning restore RS0026
            if (scaleX < 0) {
                throw new ArgumentOutOfRangeException(nameof(scaleX));
            }
            if (scaleY < 0) {
                throw new ArgumentOutOfRangeException(nameof(scaleY));
            }

            long newWidth = (long)Math.Round(Width.ToEmus() * scaleX);
            long newHeight = (long)Math.Round(Height.ToEmus() * scaleY);
            Resize(OfficeLength.Emus(newWidth), OfficeLength.Emus(newHeight), anchor);
            return this;
        }

        /// <summary>
        ///     Moves the shape by the specified offsets with explicit lengths.
        /// </summary>
        public PowerPointShape MoveBy(OfficeLength offsetX, OfficeLength offsetY) {
            Left += offsetX;
            Top += offsetY;
            return this;
        }

        // The children of p:sp, p:cxnSp and p:pic before and after their shape properties, in sequence order.
        private static readonly OfficeXmlChildOrder ConnectorOrder = new(
            new[] { OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML }, new[] { "nvCxnSpPr", "spPr", "style", "extLst" });
        private static readonly OfficeXmlChildOrder PictureOrder = new(
            new[] { OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML }, new[] { "nvPicPr", "blipFill", "spPr", "style", "extLst" });
        private static readonly OfficeXmlChildOrder GraphicFrameOrder = new(
            new[] { OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.PresentationML }, new[] { "nvGraphicFramePr", "xfrm", "graphic", "extLst" });
        private static readonly OfficeXmlChildOrder GroupOrder = new(
            new[] { OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML }, new[] { "nvGrpSpPr", "grpSpPr" });
        private static readonly OfficeXmlChildOrder TransformOrder = new(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "off", "ext", "chOff", "chExt" });

        // The shape's transform: a:xfrm in its shape properties (p:xfrm for a graphic frame). Reading must not write an
        // empty transform: a placeholder without one inherits its layout's geometry, and an empty a:off is invalid.
        private PowerPointXmlNode TransformElement(bool create) {
            PowerPointXmlNode element = XmlElement;
            if (element.NamespaceId != OfficeXmlNamespaces.PresentationML) {
                return create ? throw new NotSupportedException() : default;
            }
            switch (element.LocalName) {
                case "sp":
                case "cxnSp":
                case "pic": {
                    OfficeXmlChildOrder order = element.LocalName == "sp" ? PShapeView.Order : element.LocalName == "cxnSp" ? ConnectorOrder : PictureOrder;
                    PowerPointXmlNode properties = create ? element.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "spPr", order) : element.Child(OfficeXmlNamespaces.PresentationML, "spPr");
                    return create ? properties.GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "xfrm") : properties.Child(OfficeXmlNamespaces.DrawingML, "xfrm");
                }
                case "graphicFrame":
                    return create ? element.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "xfrm", GraphicFrameOrder) : element.Child(OfficeXmlNamespaces.PresentationML, "xfrm");
                case "grpSp": {
                    PowerPointXmlNode properties = create ? element.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "grpSpPr", GroupOrder) : element.Child(OfficeXmlNamespaces.PresentationML, "grpSpPr");
                    if (!create) return properties.Child(OfficeXmlNamespaces.DrawingML, "xfrm");
                    // A group transform carries its child offset and extents as well.
                    PowerPointXmlNode transform = properties.GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "xfrm");
                    foreach (string point in new[] { "off", "ext", "chOff", "chExt" }) EnsurePoint(transform, point);
                    return transform;
                }
                default:
                    return create ? throw new NotSupportedException() : default;
            }
        }

        // a:off, a:ext, a:chOff or a:chExt of a transform, added with zero values when missing.
        private static PowerPointXmlNode EnsurePoint(PowerPointXmlNode transform, string point) {
            PowerPointXmlNode existing = transform.Child(OfficeXmlNamespaces.DrawingML, point);
            if (!existing.IsNone) return existing;
            PowerPointXmlNode created = transform.AddOrdered(OfficeXmlNamespaces.DrawingML, point, TransformOrder);
            bool extents = point is "ext" or "chExt";
            created.SetAttribute(extents ? "cx" : "x", "0");
            created.SetAttribute(extents ? "cy" : "y", "0");
            return created;
        }

        private PowerPointXmlNode TransformPoint(string point, bool create) {
            PowerPointXmlNode transform = TransformElement(create);
            return create ? EnsurePoint(transform, point) : transform.Child(OfficeXmlNamespaces.DrawingML, point);
        }

        private void SetTransformValue(string point, string attribute, long value) {
            TransformPoint(point, create: true).SetInt64Attribute(attribute, value);
            OwnerPartRef.MarkEdited();
        }

        private bool TryGetTransformValues(out long x, out long y, out long cx, out long cy) {
            x = y = cx = cy = 0L;
            PowerPointXmlNode transform = TransformElement(create: false);
            PowerPointXmlNode offset = transform.Child(OfficeXmlNamespaces.DrawingML, "off"), extents = transform.Child(OfficeXmlNamespaces.DrawingML, "ext");
            if (offset.Int64Attribute("x") is not long offsetX || offset.Int64Attribute("y") is not long offsetY
                || extents.Int64Attribute("cx") is not long width || extents.Int64Attribute("cy") is not long height) {
                return false;
            }
            x = offsetX;
            y = offsetY;
            cx = width;
            cy = height;
            return true;
        }

        private static (double x, double y) ResolveAnchorPoint(PowerPointLayoutBox bounds, PowerPointShapeAnchor anchor) {
            double left = bounds.Left.ToEmus();
            double top = bounds.Top.ToEmus();
            double right = bounds.Right.ToEmus();
            double bottom = bounds.Bottom.ToEmus();
            double centerX = bounds.Left.ToEmus() + bounds.Width.ToEmus() / 2d;
            double centerY = bounds.Top.ToEmus() + bounds.Height.ToEmus() / 2d;

            return anchor switch {
                PowerPointShapeAnchor.TopLeft => (left, top),
                PowerPointShapeAnchor.Top => (centerX, top),
                PowerPointShapeAnchor.TopRight => (right, top),
                PowerPointShapeAnchor.Left => (left, centerY),
                PowerPointShapeAnchor.Center => (centerX, centerY),
                PowerPointShapeAnchor.Right => (right, centerY),
                PowerPointShapeAnchor.BottomLeft => (left, bottom),
                PowerPointShapeAnchor.Bottom => (centerX, bottom),
                PowerPointShapeAnchor.BottomRight => (right, bottom),
                _ => (left, top)
            };
        }

        private static (double offsetX, double offsetY) ResolveAnchorOffset(long width, long height, PowerPointShapeAnchor anchor) {
            double w = width;
            double h = height;

            return anchor switch {
                PowerPointShapeAnchor.TopLeft => (0d, 0d),
                PowerPointShapeAnchor.Top => (w / 2d, 0d),
                PowerPointShapeAnchor.TopRight => (w, 0d),
                PowerPointShapeAnchor.Left => (0d, h / 2d),
                PowerPointShapeAnchor.Center => (w / 2d, h / 2d),
                PowerPointShapeAnchor.Right => (w, h / 2d),
                PowerPointShapeAnchor.BottomLeft => (0d, h),
                PowerPointShapeAnchor.Bottom => (w / 2d, h),
                PowerPointShapeAnchor.BottomRight => (w, h),
                _ => (0d, 0d)
            };
        }
    }
}
