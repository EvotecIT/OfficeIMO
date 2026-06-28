using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Horizontal position of the shape in EMUs.
        /// </summary>
        public long Left {
            get => GetOffset().X?.Value ?? 0L;
            set => GetOffset().X = value;
        }

        /// <summary>
        ///     Horizontal position of the shape in points.
        /// </summary>
        public double LeftPoints {
            get => PowerPointUnits.ToPoints(Left);
            set => Left = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Horizontal position of the shape in centimeters.
        /// </summary>
        public double LeftCm {
            get => PowerPointUnits.ToCentimeters(Left);
            set => Left = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Horizontal position of the shape in inches.
        /// </summary>
        public double LeftInches {
            get => PowerPointUnits.ToInches(Left);
            set => Left = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Vertical position of the shape in EMUs.
        /// </summary>
        public long Top {
            get => GetOffset().Y?.Value ?? 0L;
            set => GetOffset().Y = value;
        }

        /// <summary>
        ///     Vertical position of the shape in points.
        /// </summary>
        public double TopPoints {
            get => PowerPointUnits.ToPoints(Top);
            set => Top = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Vertical position of the shape in centimeters.
        /// </summary>
        public double TopCm {
            get => PowerPointUnits.ToCentimeters(Top);
            set => Top = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Vertical position of the shape in inches.
        /// </summary>
        public double TopInches {
            get => PowerPointUnits.ToInches(Top);
            set => Top = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Width of the shape in EMUs.
        /// </summary>
        public long Width {
            get => GetExtents().Cx?.Value ?? 0L;
            set => GetExtents().Cx = value;
        }

        /// <summary>
        ///     Width of the shape in points.
        /// </summary>
        public double WidthPoints {
            get => PowerPointUnits.ToPoints(Width);
            set => Width = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Width of the shape in centimeters.
        /// </summary>
        public double WidthCm {
            get => PowerPointUnits.ToCentimeters(Width);
            set => Width = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Width of the shape in inches.
        /// </summary>
        public double WidthInches {
            get => PowerPointUnits.ToInches(Width);
            set => Width = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Height of the shape in EMUs.
        /// </summary>
        public long Height {
            get => GetExtents().Cy?.Value ?? 0L;
            set => GetExtents().Cy = value;
        }

        /// <summary>
        ///     Height of the shape in points.
        /// </summary>
        public double HeightPoints {
            get => PowerPointUnits.ToPoints(Height);
            set => Height = PowerPointUnits.FromPoints(value);
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

        /// <summary>
        ///     Height of the shape in centimeters.
        /// </summary>
        public double HeightCm {
            get => PowerPointUnits.ToCentimeters(Height);
            set => Height = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Height of the shape in inches.
        /// </summary>
        public double HeightInches {
            get => PowerPointUnits.ToInches(Height);
            set => Height = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Right position of the shape in EMUs.
        /// </summary>
        public long Right {
            get => Left + Width;
            set => Left = value - Width;
        }

        /// <summary>
        ///     Right position of the shape in points.
        /// </summary>
        public double RightPoints {
            get => PowerPointUnits.ToPoints(Right);
            set => Right = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Right position of the shape in centimeters.
        /// </summary>
        public double RightCm {
            get => PowerPointUnits.ToCentimeters(Right);
            set => Right = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Right position of the shape in inches.
        /// </summary>
        public double RightInches {
            get => PowerPointUnits.ToInches(Right);
            set => Right = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Bottom position of the shape in EMUs.
        /// </summary>
        public long Bottom {
            get => Top + Height;
            set => Top = value - Height;
        }

        /// <summary>
        ///     Bottom position of the shape in points.
        /// </summary>
        public double BottomPoints {
            get => PowerPointUnits.ToPoints(Bottom);
            set => Bottom = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Bottom position of the shape in centimeters.
        /// </summary>
        public double BottomCm {
            get => PowerPointUnits.ToCentimeters(Bottom);
            set => Bottom = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Bottom position of the shape in inches.
        /// </summary>
        public double BottomInches {
            get => PowerPointUnits.ToInches(Bottom);
            set => Bottom = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Horizontal center of the shape in EMUs.
        /// </summary>
        public long CenterX {
            get => Left + (long)Math.Round(Width / 2d);
            set => Left = (long)Math.Round(value - Width / 2d);
        }

        /// <summary>
        ///     Horizontal center of the shape in points.
        /// </summary>
        public double CenterXPoints {
            get => PowerPointUnits.ToPoints(CenterX);
            set => CenterX = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Horizontal center of the shape in centimeters.
        /// </summary>
        public double CenterXCm {
            get => PowerPointUnits.ToCentimeters(CenterX);
            set => CenterX = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Horizontal center of the shape in inches.
        /// </summary>
        public double CenterXInches {
            get => PowerPointUnits.ToInches(CenterX);
            set => CenterX = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Vertical center of the shape in EMUs.
        /// </summary>
        public long CenterY {
            get => Top + (long)Math.Round(Height / 2d);
            set => Top = (long)Math.Round(value - Height / 2d);
        }

        /// <summary>
        ///     Vertical center of the shape in points.
        /// </summary>
        public double CenterYPoints {
            get => PowerPointUnits.ToPoints(CenterY);
            set => CenterY = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Vertical center of the shape in centimeters.
        /// </summary>
        public double CenterYCm {
            get => PowerPointUnits.ToCentimeters(CenterY);
            set => CenterY = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Vertical center of the shape in inches.
        /// </summary>
        public double CenterYInches {
            get => PowerPointUnits.ToInches(CenterY);
            set => CenterY = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Gets or sets the shape bounds in EMUs.
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
        ///     Sets position using points.
        /// </summary>
        public void SetPositionPoints(double leftPoints, double topPoints) {
            LeftPoints = leftPoints;
            TopPoints = topPoints;
        }

        /// <summary>
        ///     Sets position using centimeters.
        /// </summary>
        public void SetPositionCm(double leftCm, double topCm) {
            LeftCm = leftCm;
            TopCm = topCm;
        }

        /// <summary>
        ///     Sets position using inches.
        /// </summary>
        public void SetPositionInches(double leftInches, double topInches) {
            LeftInches = leftInches;
            TopInches = topInches;
        }

        /// <summary>
        ///     Sets size using points.
        /// </summary>
        public void SetSizePoints(double widthPoints, double heightPoints) {
            WidthPoints = widthPoints;
            HeightPoints = heightPoints;
        }

        /// <summary>
        ///     Sets size using centimeters.
        /// </summary>
        public void SetSizeCm(double widthCm, double heightCm) {
            WidthCm = widthCm;
            HeightCm = heightCm;
        }

        /// <summary>
        ///     Sets size using inches.
        /// </summary>
        public void SetSizeInches(double widthInches, double heightInches) {
            WidthInches = widthInches;
            HeightInches = heightInches;
        }

        /// <summary>
        ///     Resizes the shape while keeping the specified anchor fixed.
        /// </summary>
        public void Resize(long width, long height, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            if (width < 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height < 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            PowerPointLayoutBox bounds = Bounds;
            (double anchorX, double anchorY) = ResolveAnchorPoint(bounds, anchor);
            (double offsetX, double offsetY) = ResolveAnchorOffset(width, height, anchor);

            Width = width;
            Height = height;
            Left = (long)Math.Round(anchorX - offsetX);
            Top = (long)Math.Round(anchorY - offsetY);
        }

        /// <summary>
        ///     Resizes the shape using points while keeping the specified anchor fixed.
        /// </summary>
        public void ResizePoints(double widthPoints, double heightPoints, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            Resize(PowerPointUnits.FromPoints(widthPoints), PowerPointUnits.FromPoints(heightPoints), anchor);
        }

        /// <summary>
        ///     Resizes the shape using centimeters while keeping the specified anchor fixed.
        /// </summary>
        public void ResizeCm(double widthCm, double heightCm, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            Resize(PowerPointUnits.FromCentimeters(widthCm), PowerPointUnits.FromCentimeters(heightCm), anchor);
        }

        /// <summary>
        ///     Resizes the shape using inches while keeping the specified anchor fixed.
        /// </summary>
        public void ResizeInches(double widthInches, double heightInches, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.TopLeft) {
            Resize(PowerPointUnits.FromInches(widthInches), PowerPointUnits.FromInches(heightInches), anchor);
        }

        /// <summary>
        ///     Scales the shape uniformly using the specified anchor.
        /// </summary>
        public void Scale(double scale, PowerPointShapeAnchor anchor = PowerPointShapeAnchor.Center) {
            if (scale < 0) {
                throw new ArgumentOutOfRangeException(nameof(scale));
            }

            long newWidth = (long)Math.Round(Width * scale);
            long newHeight = (long)Math.Round(Height * scale);
            Resize(newWidth, newHeight, anchor);
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

            long newWidth = (long)Math.Round(Width * scaleX);
            long newHeight = (long)Math.Round(Height * scaleY);
            Resize(newWidth, newHeight, anchor);
        }

        /// <summary>
        ///     Moves the shape by the specified offsets in EMUs.
        /// </summary>
        public void MoveBy(long offsetX, long offsetY) {
            Left += offsetX;
            Top += offsetY;
        }

        /// <summary>
        ///     Moves the shape by the specified offsets in points.
        /// </summary>
        public void MoveByPoints(double offsetXPoints, double offsetYPoints) {
            MoveBy(PowerPointUnits.FromPoints(offsetXPoints), PowerPointUnits.FromPoints(offsetYPoints));
        }

        /// <summary>
        ///     Moves the shape by the specified offsets in centimeters.
        /// </summary>
        public void MoveByCm(double offsetXCm, double offsetYCm) {
            MoveBy(PowerPointUnits.FromCentimeters(offsetXCm), PowerPointUnits.FromCentimeters(offsetYCm));
        }

        /// <summary>
        ///     Moves the shape by the specified offsets in inches.
        /// </summary>
        public void MoveByInches(double offsetXInches, double offsetYInches) {
            MoveBy(PowerPointUnits.FromInches(offsetXInches), PowerPointUnits.FromInches(offsetYInches));
        }

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
            double left = bounds.Left;
            double top = bounds.Top;
            double right = bounds.Right;
            double bottom = bounds.Bottom;
            double centerX = bounds.Left + bounds.Width / 2d;
            double centerY = bounds.Top + bounds.Height / 2d;

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
