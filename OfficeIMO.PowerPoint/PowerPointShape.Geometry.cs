using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Left of the shape.</summary>
        public OfficeLength Left {
            get => OfficeLength.Emus(PeekOffset()?.X?.Value ?? 0L);
            set => GetOffset().X = value.ToEmus();
        }
        /// <summary>Top of the shape.</summary>
        public OfficeLength Top {
            get => OfficeLength.Emus(PeekOffset()?.Y?.Value ?? 0L);
            set => GetOffset().Y = value.ToEmus();
        }
        /// <summary>Width of the shape.</summary>
        public OfficeLength Width {
            get => OfficeLength.Emus(PeekExtents()?.Cx?.Value ?? 0L);
            set => GetExtents().Cx = value.ToEmus();
        }
        /// <summary>Height of the shape.</summary>
        public OfficeLength Height {
            get => OfficeLength.Emus(PeekExtents()?.Cy?.Value ?? 0L);
            set => GetExtents().Cy = value.ToEmus();
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
        public void SetPosition(OfficeLength left, OfficeLength top) {
            Left = left;
            Top = top;
        }

        /// <summary>
        ///     Sets size with explicit lengths.
        /// </summary>
        public void SetSize(OfficeLength width, OfficeLength height) {
            Width = width;
            Height = height;
        }

        /// <summary>
        ///     Resizes the shape while keeping the specified anchor fixed.
        /// </summary>
        public void Resize(OfficeLength width, OfficeLength height, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            long widthValue = width.ToEmus();
            long heightValue = height.ToEmus();
            if (widthValue < 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (heightValue < 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            PowerPointLayoutBox bounds = Bounds;
            (double anchorX, double anchorY) = ResolveAnchorPoint(bounds, anchor);
            (double offsetX, double offsetY) = ResolveAnchorOffset(widthValue, heightValue, anchor);

            Width = OfficeLength.Emus(widthValue);
            Height = OfficeLength.Emus(heightValue);
            Left = OfficeLength.Emus((long)Math.Round(anchorX - offsetX));
            Top = OfficeLength.Emus((long)Math.Round(anchorY - offsetY));
        }

        /// <summary>
        ///     Scales the shape uniformly using the specified anchor.
        /// </summary>
        public void Scale(double scale, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.Center) {
            if (scale < 0) {
                throw new ArgumentOutOfRangeException(nameof(scale));
            }

            long newWidth = (long)Math.Round(Width.ToEmus() * scale);
            long newHeight = (long)Math.Round(Height.ToEmus() * scale);
            Resize(OfficeLength.Emus(newWidth), OfficeLength.Emus(newHeight), anchor);
        }

        /// <summary>
        ///     Scales the shape using separate X/Y factors and the specified anchor.
        /// </summary>
        public void Scale(double scaleX, double scaleY, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.Center) {
            if (scaleX < 0) {
                throw new ArgumentOutOfRangeException(nameof(scaleX));
            }
            if (scaleY < 0) {
                throw new ArgumentOutOfRangeException(nameof(scaleY));
            }

            long newWidth = (long)Math.Round(Width.ToEmus() * scaleX);
            long newHeight = (long)Math.Round(Height.ToEmus() * scaleY);
            Resize(OfficeLength.Emus(newWidth), OfficeLength.Emus(newHeight), anchor);
        }

        /// <summary>
        ///     Moves the shape by the specified offsets with explicit lengths.
        /// </summary>
        public void MoveBy(OfficeLength offsetX, OfficeLength offsetY) {
            Left += offsetX;
            Top += offsetY;
        }

        // Reading a position must not write an empty transform: a placeholder without one inherits its layout's
        // geometry, and an empty <a:off/> is invalid.
        private A.Offset? PeekOffset() => Element switch {
            Shape s => s.ShapeProperties?.Transform2D?.Offset,
            ConnectionShape c => c.ShapeProperties?.Transform2D?.Offset,
            Picture p => p.ShapeProperties?.Transform2D?.Offset,
            GraphicFrame g => g.Transform?.Offset,
            GroupShape g => g.GroupShapeProperties?.TransformGroup?.Offset,
            _ => GetOffset()
        };

        private A.Extents? PeekExtents() => Element switch {
            Shape s => s.ShapeProperties?.Transform2D?.Extents,
            ConnectionShape c => c.ShapeProperties?.Transform2D?.Extents,
            Picture p => p.ShapeProperties?.Transform2D?.Extents,
            GraphicFrame g => g.Transform?.Extents,
            GroupShape g => g.GroupShapeProperties?.TransformGroup?.Extents,
            _ => GetExtents()
        };

        private A.Offset GetOffset() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return s.ShapeProperties.Transform2D.Offset;
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return c.ShapeProperties.Transform2D.Offset;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return p.ShapeProperties.Transform2D.Offset;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Offset ??= new A.Offset();
                    return g.Transform.Offset;
                case GroupShape g:
                    A.TransformGroup transform = EnsureTransformGroup(g);
                    transform.Offset ??= new A.Offset();
                    return transform.Offset;
                default:
                    throw new NotSupportedException();
            }
        }

        private A.Extents GetExtents() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return s.ShapeProperties.Transform2D.Extents;
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return c.ShapeProperties.Transform2D.Extents;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return p.ShapeProperties.Transform2D.Extents;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Extents ??= new A.Extents();
                    return g.Transform.Extents;
                case GroupShape g:
                    A.TransformGroup transform = EnsureTransformGroup(g);
                    transform.Extents ??= new A.Extents();
                    return transform.Extents;
                default:
                    throw new NotSupportedException();
            }
        }

        private bool TryGetTransformValues(out long x, out long y, out long cx, out long cy) {
            x = 0L;
            y = 0L;
            cx = 0L;
            cy = 0L;
            (A.Offset? offset, A.Extents? extents) = Element switch {
                Shape s => (s.ShapeProperties?.Transform2D?.Offset, s.ShapeProperties?.Transform2D?.Extents),
                ConnectionShape c => (c.ShapeProperties?.Transform2D?.Offset, c.ShapeProperties?.Transform2D?.Extents),
                Picture p => (p.ShapeProperties?.Transform2D?.Offset, p.ShapeProperties?.Transform2D?.Extents),
                GraphicFrame g => (g.Transform?.Offset, g.Transform?.Extents),
                GroupShape g => (g.GroupShapeProperties?.TransformGroup?.Offset, g.GroupShapeProperties?.TransformGroup?.Extents),
                _ => (null, null)
            };

            long? offsetX = offset?.X?.Value;
            long? offsetY = offset?.Y?.Value;
            long? width = extents?.Cx?.Value;
            long? height = extents?.Cy?.Value;
            if (!offsetX.HasValue || !offsetY.HasValue || !width.HasValue || !height.HasValue) {
                return false;
            }

            x = offsetX.Value;
            y = offsetY.Value;
            cx = width.Value;
            cy = height.Value;
            return true;
        }

        private static A.TransformGroup EnsureTransformGroup(GroupShape group) {
            group.GroupShapeProperties ??= new GroupShapeProperties();
            A.TransformGroup transform = group.GroupShapeProperties.TransformGroup ??= new A.TransformGroup();
            transform.Offset ??= new A.Offset();
            transform.Extents ??= new A.Extents();
            transform.ChildOffset ??= new A.ChildOffset();
            transform.ChildExtents ??= new A.ChildExtents();
            return transform;
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
