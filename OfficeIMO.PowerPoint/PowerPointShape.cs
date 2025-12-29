using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract class PowerPointShape {
        private const int EmusPerPoint = 12700;

        internal PowerPointShape(OpenXmlElement element) {
            Element = element;
        }

        internal OpenXmlElement Element { get; }

        /// <summary>
        ///     Name assigned to the shape.
        /// </summary>
        public string? Name {
            get {
                switch (Element) {
                    case Shape s:
                        return s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                    case Picture p:
                        return p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                    case GraphicFrame g:
                        return g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                    case GroupShape g:
                        return g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                    default:
                        return null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the fill color of the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? FillColor {
            get {
                ShapeProperties? props = GetShapeProperties();
                A.SolidFill? solid = props?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                ShapeProperties? props = GetShapeProperties(create: value != null);
                if (props == null) {
                    return;
                }

                props.RemoveAllChildren<A.SolidFill>();
                if (value != null) {
                    props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }
            }
        }

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

        /// <summary>
        ///     Gets or sets the outline color for the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? OutlineColor {
            get {
                A.Outline? outline = GetOutline();
                return outline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            }
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.Remove();
                    return;
                }

                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
            }
        }

        /// <summary>
        ///     Gets or sets the outline width in points.
        /// </summary>
        public double? OutlineWidthPoints {
            get {
                A.Outline? outline = GetOutline();
                int? width = outline?.Width?.Value;
                return width != null ? width.Value / (double)EmusPerPoint : null;
            }
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.Width = null;
                    if (!outline.HasChildren) {
                        outline.Remove();
                    }
                    return;
                }

                outline.Width = (int)Math.Round(value.Value * EmusPerPoint);
            }
        }

        /// <summary>
        ///     Gets or sets the outline dash preset.
        /// </summary>
        public A.PresetLineDashValues? OutlineDash {
            get => GetOutline()?.GetFirstChild<A.PresetDash>()?.Val?.Value;
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.GetFirstChild<A.PresetDash>()?.Remove();
                    return;
                }

                A.PresetDash dash = outline.GetFirstChild<A.PresetDash>() ?? new A.PresetDash();
                dash.Val = value.Value;
                if (dash.Parent == null) {
                    outline.Append(dash);
                }
            }
        }

        /// <summary>
        ///     Sets arrowheads for line-based shapes.
        /// </summary>
        public void SetLineEnds(A.LineEndValues? startType, A.LineEndValues? endType, A.LineEndWidthValues? width = null, A.LineEndLengthValues? length = null) {
            bool create = startType != null || endType != null || width != null || length != null;
            A.Outline? outline = GetOutline(create: create);
            if (outline == null) {
                return;
            }

            ApplyLineEnd(outline, startType, width, length, isStart: true);
            ApplyLineEnd(outline, endType, width, length, isStart: false);
        }

        /// <summary>
        ///     Applies a drop shadow to the shape.
        /// </summary>
        public void SetShadow(string color, double blurPoints = 4, double distancePoints = 3,
            double angleDegrees = 45, int transparencyPercent = 35, bool rotateWithShape = false) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Shadow color cannot be null or empty.", nameof(color));
            }
            if (blurPoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(blurPoints));
            }
            if (distancePoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(distancePoints));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            A.OuterShadow? shadow = GetOuterShadow(create: true);
            if (shadow == null) {
                return;
            }

            shadow.BlurRadius = PowerPointUnits.FromPoints(blurPoints);
            shadow.Distance = PowerPointUnits.FromPoints(distancePoints);
            shadow.Direction = ToShadowAngle(angleDegrees);
            shadow.RotateWithShape = rotateWithShape;

            RemoveShadowColors(shadow);
            A.RgbColorModelHex rgb = new() { Val = color };
            int alpha = 100000 - transparencyPercent * 1000;
            rgb.Append(new A.Alpha { Val = alpha });
            shadow.Append(rgb);
        }

        /// <summary>
        ///     Removes any drop shadow from the shape.
        /// </summary>
        public void ClearShadow() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.OuterShadow>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a glow effect to the shape.
        /// </summary>
        public void SetGlow(string color, double radiusPoints = 4, int transparencyPercent = 30) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Glow color cannot be null or empty.", nameof(color));
            }
            if (radiusPoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(radiusPoints));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            A.Glow? glow = GetGlow(create: true);
            if (glow == null) {
                return;
            }

            glow.Radius = PowerPointUnits.FromPoints(radiusPoints);
            RemoveGlowColors(glow);

            A.RgbColorModelHex rgb = new() { Val = color };
            int alpha = 100000 - transparencyPercent * 1000;
            rgb.Append(new A.Alpha { Val = alpha });
            glow.Append(rgb);
        }

        /// <summary>
        ///     Removes any glow effect from the shape.
        /// </summary>
        public void ClearGlow() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Glow>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a soft edges effect to the shape.
        /// </summary>
        public void SetSoftEdges(double radiusPoints) {
            if (radiusPoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(radiusPoints));
            }

            A.SoftEdge? softEdge = GetSoftEdge(create: true);
            if (softEdge == null) {
                return;
            }

            softEdge.Radius = PowerPointUnits.FromPoints(radiusPoints);
        }

        /// <summary>
        ///     Removes any soft edges effect from the shape.
        /// </summary>
        public void ClearSoftEdges() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.SoftEdge>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a blur effect to the shape.
        /// </summary>
        public void SetBlur(double radiusPoints, bool grow = false) {
            if (radiusPoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(radiusPoints));
            }

            A.Blur? blur = GetBlur(create: true);
            if (blur == null) {
                return;
            }

            blur.Radius = PowerPointUnits.FromPoints(radiusPoints);
            blur.Grow = grow;
        }

        /// <summary>
        ///     Removes any blur effect from the shape.
        /// </summary>
        public void ClearBlur() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Blur>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a reflection effect to the shape.
        /// </summary>
        public void SetReflection(
            double blurPoints = 4,
            double distancePoints = 2,
            double directionDegrees = 270,
            double fadeDirectionDegrees = 90,
            int startOpacityPercent = 50,
            int endOpacityPercent = 0,
            int startPositionPercent = 0,
            int endPositionPercent = 100,
            A.RectangleAlignmentValues? alignment = null,
            bool rotateWithShape = false) {
            if (blurPoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(blurPoints));
            }
            if (distancePoints < 0) {
                throw new ArgumentOutOfRangeException(nameof(distancePoints));
            }
            if (startOpacityPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(startOpacityPercent), "Opacity must be between 0 and 100.");
            }
            if (endOpacityPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(endOpacityPercent), "Opacity must be between 0 and 100.");
            }
            if (startPositionPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(startPositionPercent), "Position must be between 0 and 100.");
            }
            if (endPositionPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(endPositionPercent), "Position must be between 0 and 100.");
            }

            A.Reflection? reflection = GetReflection(create: true);
            if (reflection == null) {
                return;
            }

            reflection.BlurRadius = PowerPointUnits.FromPoints(blurPoints);
            reflection.Distance = PowerPointUnits.FromPoints(distancePoints);
            reflection.Direction = ToShadowAngle(directionDegrees);
            reflection.FadeDirection = ToShadowAngle(fadeDirectionDegrees);
            reflection.StartOpacity = ToAlphaValue(startOpacityPercent);
            reflection.EndAlpha = ToAlphaValue(endOpacityPercent);
            reflection.StartPosition = ToPercentValue(startPositionPercent);
            reflection.EndPosition = ToPercentValue(endPositionPercent);
            reflection.Alignment = alignment ?? A.RectangleAlignmentValues.Bottom;
            reflection.RotateWithShape = rotateWithShape;
        }

        /// <summary>
        ///     Removes any reflection effect from the shape.
        /// </summary>
        public void ClearReflection() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Reflection>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Gets or sets the fill transparency percentage (0-100). 0 = opaque, 100 = fully transparent.
        /// </summary>
        public int? FillTransparency {
            get {
                ShapeProperties? props = GetShapeProperties();
                A.SolidFill? solid = props?.GetFirstChild<A.SolidFill>();
                A.RgbColorModelHex? color = solid?.RgbColorModelHex;
                A.Alpha? alpha = color?.GetFirstChild<A.Alpha>();
                int? val = alpha?.Val?.Value;
                if (val == null) {
                    return null;
                }
                return (int)Math.Round((100000 - val.Value) / 1000d);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }

                ShapeProperties? props = GetShapeProperties(create: value != null);
                if (props == null) {
                    return;
                }

                A.SolidFill? solid = props.GetFirstChild<A.SolidFill>();
                if (solid == null) {
                    if (value == null) {
                        return;
                    }
                    solid = new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" });
                    props.Append(solid);
                }

                A.RgbColorModelHex? rgb = solid.RgbColorModelHex ?? new A.RgbColorModelHex { Val = "FFFFFF" };
                solid.RgbColorModelHex ??= rgb;
                A.Alpha? alpha = rgb.GetFirstChild<A.Alpha>();
                if (value == null) {
                    alpha?.Remove();
                    return;
                }

                if (alpha == null) {
                    alpha = new A.Alpha();
                    rgb.Append(alpha);
                }
                alpha.Val = 100000 - value.Value * 1000;
            }
        }

        /// <summary>
        ///     Gets or sets rotation in degrees.
        /// </summary>
        public double? Rotation {
            get {
                int? rotation = GetRotation();
                return rotation != null ? rotation.Value / 60000d : null;
            }
            set {
                int? rotation = value != null ? (int)Math.Round(value.Value * 60000d) : null;
                SetRotation(rotation);
            }
        }

        /// <summary>
        ///     Gets or sets horizontal flip.
        /// </summary>
        public bool? HorizontalFlip {
            get => GetHorizontalFlip();
            set => SetHorizontalFlip(value);
        }

        /// <summary>
        ///     Gets or sets vertical flip.
        /// </summary>
        public bool? VerticalFlip {
            get => GetVerticalFlip();
            set => SetVerticalFlip(value);
        }

        /// <summary>
        ///     Moves the shape to the front (top) of the z-order within its parent.
        /// </summary>
        public void BringToFront() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();
            parent.Append(Element);
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order within its parent.
        /// </summary>
        public void SendToBack() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();

            OpenXmlElement? insertBefore = null;
            foreach (OpenXmlElement child in parent.ChildElements) {
                if (child is NonVisualGroupShapeProperties || child is GroupShapeProperties) {
                    continue;
                }
                insertBefore = child;
                break;
            }

            if (insertBefore != null) {
                parent.InsertBefore(Element, insertBefore);
            } else {
                parent.Append(Element);
            }
        }

        private A.Offset GetOffset() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return s.ShapeProperties.Transform2D.Offset;
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

        private ShapeProperties? GetShapeProperties(bool create = false) {
            switch (Element) {
                case Shape s:
                    if (create) {
                        s.ShapeProperties ??= new ShapeProperties();
                    }
                    return s.ShapeProperties;
                case Picture p:
                    if (create) {
                        p.ShapeProperties ??= new ShapeProperties();
                    }
                    return p.ShapeProperties;
                default:
                    return null;
            }
        }

        private A.Outline? GetOutline(bool create = false) {
            ShapeProperties? props = GetShapeProperties(create);
            if (props == null) {
                return null;
            }

            A.Outline? outline = props.GetFirstChild<A.Outline>();
            if (outline == null && create) {
                outline = new A.Outline();
                props.Append(outline);
            }

            return outline;
        }

        private A.EffectList? GetEffectList(bool create = false) {
            ShapeProperties? props = GetShapeProperties(create);
            if (props == null) {
                return null;
            }

            A.EffectList? effectList = props.GetFirstChild<A.EffectList>();
            if (effectList == null && create) {
                effectList = new A.EffectList();
                props.Append(effectList);
            }

            return effectList;
        }

        private A.OuterShadow? GetOuterShadow(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.OuterShadow? shadow = effects.GetFirstChild<A.OuterShadow>();
            if (shadow == null && create) {
                shadow = new A.OuterShadow();
                effects.Append(shadow);
            }

            return shadow;
        }

        private A.Blur? GetBlur(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Blur? blur = effects.GetFirstChild<A.Blur>();
            if (blur == null && create) {
                blur = new A.Blur();
                effects.Append(blur);
            }

            return blur;
        }

        private A.Reflection? GetReflection(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Reflection? reflection = effects.GetFirstChild<A.Reflection>();
            if (reflection == null && create) {
                reflection = new A.Reflection();
                effects.Append(reflection);
            }

            return reflection;
        }

        private A.Glow? GetGlow(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Glow? glow = effects.GetFirstChild<A.Glow>();
            if (glow == null && create) {
                glow = new A.Glow();
                effects.Append(glow);
            }

            return glow;
        }

        private A.SoftEdge? GetSoftEdge(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.SoftEdge? softEdge = effects.GetFirstChild<A.SoftEdge>();
            if (softEdge == null && create) {
                softEdge = new A.SoftEdge();
                effects.Append(softEdge);
            }

            return softEdge;
        }

        private static void RemoveShadowColors(A.OuterShadow shadow) {
            shadow.RemoveAllChildren<A.RgbColorModelHex>();
            shadow.RemoveAllChildren<A.SchemeColor>();
            shadow.RemoveAllChildren<A.SystemColor>();
            shadow.RemoveAllChildren<A.PresetColor>();
        }

        private static void RemoveGlowColors(A.Glow glow) {
            glow.RemoveAllChildren<A.RgbColorModelHex>();
            glow.RemoveAllChildren<A.SchemeColor>();
            glow.RemoveAllChildren<A.SystemColor>();
            glow.RemoveAllChildren<A.PresetColor>();
        }

        private static int ToShadowAngle(double degrees) {
            double normalized = degrees % 360d;
            if (normalized < 0) {
                normalized += 360d;
            }
            return (int)Math.Round(normalized * 60000d);
        }

        private static int ToAlphaValue(int percent) {
            return percent * 1000;
        }

        private static int ToPercentValue(int percent) {
            return percent * 1000;
        }

        private static void ApplyLineEnd(A.Outline outline, A.LineEndValues? type, A.LineEndWidthValues? width, A.LineEndLengthValues? length, bool isStart) {
            bool hasData = type != null || width != null || length != null;
            if (isStart) {
                A.HeadEnd? head = outline.GetFirstChild<A.HeadEnd>();
                if (!hasData) {
                    head?.Remove();
                    return;
                }

                head ??= new A.HeadEnd();
                head.Type = type ?? A.LineEndValues.None;
                if (width != null) {
                    head.Width = width.Value;
                }
                if (length != null) {
                    head.Length = length.Value;
                }
                if (head.Parent == null) {
                    outline.Append(head);
                }
            } else {
                A.TailEnd? tail = outline.GetFirstChild<A.TailEnd>();
                if (!hasData) {
                    tail?.Remove();
                    return;
                }

                tail ??= new A.TailEnd();
                tail.Type = type ?? A.LineEndValues.None;
                if (width != null) {
                    tail.Width = width.Value;
                }
                if (length != null) {
                    tail.Length = length.Value;
                }
                if (tail.Parent == null) {
                    outline.Append(tail);
                }
            }
        }

        private int? GetRotation() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.Rotation?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.Rotation?.Value,
                GraphicFrame g => g.Transform?.Rotation?.Value,
                _ => null
            };
        }

        private void SetRotation(int? rotation) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Rotation = rotation;
                    break;
            }
        }

        private bool? GetHorizontalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                GraphicFrame g => g.Transform?.HorizontalFlip?.Value,
                _ => null
            };
        }

        private void SetHorizontalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.HorizontalFlip = value;
                    break;
            }
        }

        private bool? GetVerticalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                GraphicFrame g => g.Transform?.VerticalFlip?.Value,
                _ => null
            };
        }

        private void SetVerticalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.VerticalFlip = value;
                    break;
            }
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
